import { convertInchesToTwip, LevelFormat, Packer } from 'docx';
import JSZip from 'jszip';
import { Schema } from 'prosemirror-model';
import { TextDecoder, TextEncoder } from 'util';
import { DocxSerializer, DocxSerializerState, NodeSerializer } from '../src/serializer';
import { getImageType } from '../src/utils';

Object.assign(global, { TextDecoder, TextEncoder });

const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: { content: 'text*', group: 'block' },
    text: { group: 'inline' },
    image: { attrs: { src: {} }, group: 'block' },
    ordered_list: {
      attrs: { order: { default: 1 }, orderPrefix: { default: '' } },
      content: 'list_item+',
      group: 'block',
    },
    bullet_list: { content: 'list_item+', group: 'block' },
    list_item: { content: 'paragraph block*' },
    table: { content: 'table_row+', group: 'block' },
    table_row: { content: 'table_cell+' },
    table_cell: {
      attrs: { colspan: { default: 1 }, rowspan: { default: 1 } },
      content: 'block+',
    },
  },
});

const paragraph = (text: string) => schema.node('paragraph', null, schema.text(text));

const listItem = (text: string, blocks: any[] = []) =>
  schema.node('list_item', null, [paragraph(text), ...blocks]);

const orderedList = (order: number, items: any[], orderPrefix = '') =>
  schema.node('ordered_list', { order, orderPrefix }, items);

const bulletList = (items: any[]) => schema.node('bullet_list', null, items);

const table = (text: string) =>
  schema.node('table', null, [
    schema.node('table_row', null, [schema.node('table_cell', null, [paragraph(text)])]),
  ]);

const nodes: NodeSerializer = {
  text(state, node) {
    state.text(node.text || '');
  },
  paragraph(state, node) {
    state.renderInline(node);
    state.closeBlock(node);
  },
  ordered_list(state, node) {
    state.renderList(node, 'numbered');
  },
  bullet_list(state, node) {
    state.renderList(node, 'bullets');
  },
  list_item(state, node, parent, index) {
    state.renderListItem(node, index);
  },
  table_cell() {},
  table_row() {},
  table(state, node) {
    state.table(node);
  },
  image(state, node) {
    state.image(node.attrs.src);
    state.closeBlock(node);
  },
};

function render(content: any, numberingStyles: any = null, nodeSerializers = nodes) {
  const state = new DocxSerializerState(
    nodeSerializers,
    {},
    { getImageBuffer: () => null as any },
    {},
    'hr',
    numberingStyles,
  );
  state.renderContent(content);
  return state;
}

async function serializeXml(content: any, getImageBuffer: any = () => null as any) {
  const serializer = new DocxSerializer(nodes, {});
  const wordDocument = serializer.serialize(
    content,
    { getImageBuffer },
    '',
    [],
    {},
    {},
  );
  const zip = await JSZip.loadAsync(await Packer.toBuffer(wordDocument));
  return {
    document: await zip.file('word/document.xml')!.async('text'),
    numbering: await zip.file('word/numbering.xml')!.async('text'),
    files: Object.keys(zip.files),
  };
}

describe('ordered list numbering', () => {
  it('preserves starts for lists separated by tables', () => {
    const doc = schema.node('doc', null, [
      orderedList(2, [listItem('Two')]),
      table('A'),
      orderedList(3, [listItem('Three')]),
      table('B'),
      orderedList(4, [listItem('Four')]),
    ]);

    const starts = render(doc).numbering.map((numbering) => (numbering.levels[0] as any).start);
    expect(starts).toEqual([2, 3, 4]);
  });

  it('combines explicit prefixes with the selected numbering style', () => {
    const doc = schema.node('doc', null, [orderedList(3, [listItem('Three')], '1.2.')]);
    const state = render(doc, {
      numbered: {
        listStyleType: 'lower-alpha',
        listStyleColor: '2850c8',
        size: 24,
      },
    });
    const firstLevel = state.numbering[0].levels[0] as any;

    expect(firstLevel.start).toBe(3);
    expect(firstLevel.format).toBe(LevelFormat.LOWER_LETTER);
    expect(firstLevel.text).toBe('a.b.%1.');
    expect(firstLevel.style.run).toMatchObject({ color: '2850c8', size: 24 });
  });

  it('uses one native multilevel definition for compatible ordered lists', () => {
    const nested = orderedList(3, [listItem('Child three'), listItem('Child four')]);
    const doc = schema.node('doc', null, [orderedList(1, [listItem('Parent', [nested])])]);
    const state = render(doc);

    expect(state.numbering).toHaveLength(1);
    expect((state.numbering[0].levels[1] as any).start).toBe(3);
    expect((state.numbering[0].levels[1] as any).text).toBe('%1.%2.');
  });

  it('creates a separate indented definition for mixed nested list types', () => {
    const nested = bulletList([listItem('Bullet')]);
    const doc = schema.node('doc', null, [orderedList(1, [listItem('Parent', [nested])])]);
    const state = render(doc);

    expect(state.numbering).toHaveLength(2);
    expect((state.numbering[0].levels[0] as any).format).toBe(LevelFormat.DECIMAL);
    expect((state.numbering[1].levels[0] as any).format).toBe(LevelFormat.BULLET);
    expect((state.numbering[1].levels[0] as any).style.paragraph.indent.left).toBeGreaterThan(
      (state.numbering[0].levels[0] as any).style.paragraph.indent.left,
    );
  });

  it('preserves the legacy default bullet indentation', () => {
    const doc = schema.node('doc', null, [bulletList([listItem('Bullet')])]);
    const indent = (render(doc).numbering[0].levels[0] as any).style.paragraph.indent;

    expect(indent).toEqual({
      left: convertInchesToTwip(0.5),
      hanging: convertInchesToTwip(0.18),
    });
  });

  it('preserves custom bullet indentation', () => {
    const doc = schema.node('doc', null, [bulletList([listItem('Bullet')])]);
    const state = render(doc, { bullets: { listStyleType: 'circle' } });
    const indent = (state.numbering[0].levels[0] as any).style.paragraph.indent;

    expect(indent).toEqual({
      left: convertInchesToTwip(0.25),
      hanging: convertInchesToTwip(0.25),
    });
  });

  it('splits conflicting nested starts and preserves the parent prefix', () => {
    const doc = schema.node('doc', null, [
      orderedList(1, [
        listItem('Parent one', [orderedList(1, [listItem('Child one')])]),
        listItem('Parent two', [orderedList(3, [listItem('Child three')])]),
      ]),
    ]);
    const state = render(doc);

    expect(state.numbering).toHaveLength(2);
    expect((state.numbering[1].levels[0] as any).start).toBe(3);
    expect((state.numbering[1].levels[0] as any).text).toBe('2.%1.');
  });

  it('tracks parent prefixes for legacy list-item serializers without an index', () => {
    const legacyNodes: NodeSerializer = {
      ...nodes,
      list_item(state, node) {
        state.renderListItem(node);
      },
    };
    const doc = schema.node('doc', null, [
      orderedList(1, [
        listItem('Parent one', [orderedList(1, [listItem('Child one')])]),
        listItem('Parent two', [orderedList(3, [listItem('Child three')])]),
      ]),
    ]);
    const state = render(doc, null, legacyNodes);

    expect((state.numbering[1].levels[0] as any).text).toBe('2.%1.');
  });

  it('restarts separate nested lists inside the same parent item', () => {
    const doc = schema.node('doc', null, [
      orderedList(1, [
        listItem('Parent', [
          orderedList(1, [listItem('First child list')]),
          paragraph('Interruption'),
          orderedList(1, [listItem('Second child list')]),
        ]),
      ]),
    ]);
    const state = render(doc);

    expect(state.numbering).toHaveLength(2);
    expect((state.numbering[1].levels[0] as any).text).toBe('1.%1.');
  });

  it('keeps consecutive list-item paragraphs in one Word paragraph', () => {
    const item = schema.node('list_item', null, [paragraph('First'), paragraph('Second')]);
    const doc = schema.node('doc', null, [orderedList(1, [item])]);

    expect(render(doc).children).toHaveLength(1);
  });

  it('creates all nine native Word numbering levels for custom styles', () => {
    const doc = schema.node('doc', null, [orderedList(1, [listItem('One')])]);
    const state = render(doc, { numbered: { listStyleType: 'upper-roman' } });

    expect(state.numbering[0].levels).toHaveLength(9);
  });

  it('does not leak numbering state between serializer states', () => {
    const continued = schema.node('doc', null, [orderedList(4, [listItem('Four')])]);
    const fresh = schema.node('doc', null, [orderedList(1, [listItem('One')])]);

    render(continued);
    expect((render(fresh).numbering[0].levels[0] as any).start).toBeUndefined();
  });

  it('writes native starts and numbers only the first paragraph of an item', async () => {
    const twoParagraphItem = schema.node('list_item', null, [
      paragraph('First paragraph'),
      paragraph('Second paragraph'),
    ]);
    const doc = schema.node('doc', null, [
      orderedList(2, [twoParagraphItem]),
      table('A'),
      orderedList(4, [listItem('Four')]),
    ]);
    const xml = await serializeXml(doc);

    expect(xml.numbering).toMatch(/<w:startOverride w:val="2"\/>/);
    expect(xml.numbering).toMatch(/<w:startOverride w:val="4"\/>/);
    expect(xml.document.match(/<w:numPr>/g)).toHaveLength(2);
  });

  it('restores the parent numbering reference after a mixed nested list', async () => {
    const nested = bulletList([listItem('Nested bullet')]);
    const doc = schema.node('doc', null, [
      orderedList(1, [listItem('Parent one', [nested]), listItem('Parent two')]),
    ]);
    const xml = await serializeXml(doc);
    const references = [...xml.document.matchAll(/<w:numId w:val="(\d+)"\/>/g)].map(
      (match) => match[1],
    );

    expect(references).toHaveLength(3);
    expect(references[0]).toBe(references[2]);
    expect(references[1]).not.toBe(references[0]);
  });

  it('uses image bytes when an image URL has no extension', async () => {
    const src = 'lattics-local://assets/image-id';
    const doc = schema.node('doc', null, [schema.node('image', { src })]);
    const xml = await serializeXml(doc, () => ({
      arrayBuffer: new Uint8Array([0xff, 0xd8, 0xff, 0xe0]),
      width: 1,
      height: 1,
    }));

    expect(xml.files.some((file) => /word\/media\/[^/]+\.jpg$/.test(file))).toBe(true);
  });

  it('detects image signatures from ArrayBuffers and typed-array views', () => {
    const jpeg = new Uint8Array([0xff, 0xd8, 0xff, 0xe0]);

    expect(getImageType('lattics-local://assets/image-id', jpeg.buffer)).toBe('jpg');
    expect(getImageType('mismatched.png', jpeg)).toBe('jpg');
  });
});
