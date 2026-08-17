import { Schema } from 'prosemirror-model';
import { builders } from 'prosemirror-test-builder';

const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: { content: 'inline*', group: 'block' },
    text: { group: 'inline' },
    heading: { attrs: { level: { default: 1 } }, content: 'inline*', group: 'block' },
    horizontal_rule: { group: 'block' },
    list_item: { content: 'paragraph block*' },
    ordered_list: {
      attrs: { order: { default: 1 } },
      content: 'list_item+',
      group: 'block',
    },
    bullet_list: { content: 'list_item+', group: 'block' },
    code_block: { content: 'text*', group: 'block' },
    hard_break: { inline: true, group: 'inline' },
    image: {
      attrs: { src: {}, alt: { default: null } },
      inline: true,
      group: 'inline',
    },
    math: { content: 'text*', inline: true, group: 'inline' },
    equation: {
      attrs: { numbered: { default: true }, id: { default: null } },
      content: 'text*',
      group: 'block',
    },
    aside: { content: 'block+', group: 'block' },
  },
  marks: {
    em: {},
    strong: {},
    code: {},
    link: { attrs: { href: {} } },
    abbr: { attrs: { title: {} } },
  },
});

export const tnodes = builders(schema, {
  p: { nodeType: 'paragraph' },
  h1: { nodeType: 'heading', level: 1 },
  h2: { nodeType: 'heading', level: 2 },
  hr: { nodeType: 'horizontal_rule' },
  li: { nodeType: 'list_item' },
  ol: { nodeType: 'ordered_list' },
  ol3: { nodeType: 'ordered_list', order: 3 },
  ul: { nodeType: 'bullet_list' },
  pre: { nodeType: 'code_block' },
  br: { nodeType: 'hard_break' },
  img: { nodeType: 'image', src: 'img.png', alt: 'x' },
  a: { markType: 'link', href: 'https://example.com' },
  math: { nodeType: 'math' },
  equation: { nodeType: 'equation', numbered: true, id: 'eq1' },
  equationUnnumbered: { nodeType: 'equation', numbered: false, id: 'eq2' },
  abbr: { nodeType: 'abbr', title: 'Cascading Style Sheets' },
  aside: { nodeType: 'aside' },
});

export const tdoc = (...args: any[]) => (tnodes.doc as any)('', ...args);
