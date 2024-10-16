import { AlignmentType, BorderStyle, HeadingLevel, UnderlineType } from 'docx';
import { DocxSerializer, MarkSerializer, NodeSerializer } from './serializer';
import { coverColorToHex } from './utils';
import sizeTransfer from './sizeTransfer';

const LINE_TYPE: Record<string, UnderlineType> = {
  solid: UnderlineType.SINGLE,
  double: UnderlineType.DOUBLE,
  dotted: UnderlineType.DOTTED,
  dashed: UnderlineType.DASH,
  wavy: UnderlineType.WAVE,
};

const zoteroReg = /\(\[(.*?)\]\((zotero:\/\/(.*?))\)\)/g;

const colors = [
  {
    name: 'yellow',
    value: 'rgba(255, 195, 0, 0.2)',
  },
  {
    name: 'red',
    value: 'rgba(255, 90, 90, 0.18)',
  },
  {
    name: 'magenta',
    value: 'rgba(166, 125, 255, 0.15)',
  },
  {
    name: 'green',
    value: 'rgba(158, 255, 0, 0.2)',
  },
  {
    name: 'blue',
    value: 'rgba(52, 226, 216, 0.2)',
  },
  {
    name: 'darkYellow',
    value: 'rgba(255, 154, 61, 0.15)',
  },
  {
    name: 'lightGray',
    value: 'rgba(135, 135, 135, 0.2)',
  },
];
export const defaultNodes: NodeSerializer = {
  text(state, node) {
    state.text((node.text ?? '').replace(/\u200b/g, '').replace(/\u00A0/g, ' '));
  },
  paragraph(state, node) {
    if (node.attrs.footnotesHole) {
      state.children.push('[[THIS_IS_A_FOOTNOTES_HOLE]]');
    } else {
      state.renderInline(node);
      state.closeBlock(node);
    }
  },
  group_bio_citation(state, node) {
    state.bib_cite(node);
  },
  bio_citation(state, node) {
    state.bib_cite(node);
  },
  bio_display_citation(state, node) {
    const content = state.transformHtmlToNode(node.attrs.content);

    state.renderInline(content || node);
  },
  bibliography(state, node) {
    state.bibliography(node);
  },
  bibliography_display(state, node) {
    state.bibliography(node);
  },
  comment(state, node) {
    state.wrapComment(node);
  },
  heading(state, node) {
    state.renderInline(node);
    const heading = [
      HeadingLevel.HEADING_1,
      HeadingLevel.HEADING_2,
      HeadingLevel.HEADING_3,
      HeadingLevel.HEADING_4,
      HeadingLevel.HEADING_5,
      HeadingLevel.HEADING_6,
    ][node.attrs.level - 1];
    state.closeBlock(node, { heading, style: undefined });
  },
  hierarchy_title(state, node) {
    state.hierarchy_title(node);
  },
  blockquote(state, node) {
    state.renderContent(node, {
      style: 'IntenseQuote',
      // indent: { left: 250 },
      // border: { left: { style: BorderStyle.THICK, size: 20, color: 'efeff3', space: 40 } },
      // indent: { left: 250 },
      // border: { left: { style: BorderStyle.THICK_THIN_MEDIUM_GAP, size: 40 } },
    });
  },
  code_block(state, node) {
    // TODO: something for code
    // state.renderContent(node, {
    //   style: 'code',
    // });
    // state.closeBlock(node);
    state.addCodeBlock(node);
  },
  horizontal_rule(state, node) {
    state.horizontal_rule(node);
  },
  hard_break(state) {
    state.addRunOptions({ break: 1 });
  },
  footnote(state, node) {
    state.footnoteRef(node.attrs.id);
  },
  ordered_list(state, node) {
    state.renderList(node, 'numbered');
  },
  bullet_list(state, node) {
    state.renderList(node, 'bullets');
  },
  list_item(state, node) {
    state.renderListItem(node);
  },
  // Presentational
  image(state, node) {
    const { src, title = '', layout = 'center', width = 100 } = node.attrs;
    state.image(src, layout, width);
    state.closeBlock(node);
    if (title) state.addAside(title);
  },
  // Technical
  latex(state, node) {
    // state.math(getLatexFromNode(node), { inline: true });
    // state.math(node.attrs.input, { inline: true });
    state.imageInline(node.attrs.input);
  },
  blocked_latex(state, node) {
    // const { id = Date.now(), numbered } = node.attrs;
    // state.math(node.attrs.input, { inline: false, numbered, id });
    // state.closeBlock(node);
    state.imageInline(node.attrs.input);
    state.closeBlock(node, {
      alignment: AlignmentType.CENTER,
    });
  },
  link(state, node) {
    // Note, this is handled specifically in the serializer
    // Word treats links more like a Node rather than a mark
    state.openLink(node.attrs.href);
    state.renderInline(node);
    state.closeLink();
  },
  table_cell(state, node) {},
  table_header(state, node) {},
  table_row(state, node) {},
  table(state, node) {
    state.table(node);
    // console.log(state);
  },
  columns(state, node) {
    state.columns(node);
  },
  inlineCitation(state, node) {
    try {
      const { citeId, isFullCite } = node.attrs;
      if (isFullCite) {
        const text = (state.fullCiteContents[citeId] || '').replace(/\u200b/g, '');
        text.split('\n').forEach((line, index) => {
          if (index !== 0) state.addRunOptions({ break: 1 });
          state.text(line);
        });
        // state.text(text || '');
        // const images = [...text.matchAll(/LAT_IMAGE\(([^()]+)\)/g)];
        // if (!images.length) {
        //   state.text(text || '');
        // } else {
        //   let midText = text;
        //   for (let i = 0; i < images.length; i++) {
        //     const image = images[i];
        //     const [match, src] = image;
        //     const index = midText.indexOf(match);
        //     const before = midText.slice(0, index);
        //     const after = midText.slice(index + match.length);
        //     state.text(before);
        //     state.image(src);
        //     midText = after;
        //   }
        //   state.text(midText);
        // }
      } else {
        let href = '';
        try {
          if (String(citeId || '').startsWith('zotero')) {
            href = [...citeId.matchAll(zoteroReg)].at(-1)?.[2];
          }
        } catch (e) {
          console.log('eport cite error: ', e);
        }

        if (href) {
          state.openLink(href);
        }
        state.renderInline(node);
        if (href) {
          state.closeLink();
        }
      }
    } catch (e) {
      console.log('eport cite error: ', e);
    }
  },
  default(state, node) {
    if (node.isAtom || node.isLeaf) return;

    if (node.isInline) {
      state.renderInline(node);
    } else {
      state.renderContent(node);
    }
  },
};

export const defaultMarks: MarkSerializer = {
  italic() {
    return { italics: true };
  },
  bold() {
    return { bold: true };
  },
  color(state, node, mark) {
    return {
      color: mark.attrs.color ? coverColorToHex(mark.attrs.color) : '000000',
    };
  },
  font_family(state, node, mark) {
    return {
      font: mark.attrs.fontFamily || 'sans-serif',
    };
  },
  font_size(state, node, mark) {
    try {
      return {
        size: sizeTransfer(parseInt(mark.attrs.fontSize || '', 10)),
      };
    } catch (e) {
      return {};
    }
  },
  highlight(state, node, mark) {
    const target = colors.find((c) => c.value === mark.attrs.color);
    // console.log('highlight', target);
    return target
      ? {
          highlight: target.name,
        }
      : {};
  },
  abbr() {
    // TODO: abbreviation
    return {};
  },
  sub() {
    return { subScript: true };
  },
  sup() {
    return { superScript: true };
  },
  strike() {
    // doubleStrike!
    return { strike: true };
  },
  underline(_, __, mark) {
    const lineType = mark.attrs.lineType || 'solid';

    return {
      underline: {
        type: LINE_TYPE[lineType] || UnderlineType.SINGLE,
      },
    };
  },
  smallcaps() {
    return { smallCaps: true };
  },
  allcaps() {
    return { allCaps: true };
  },
};

export const defaultDocxSerializer = new DocxSerializer(defaultNodes, defaultMarks);
