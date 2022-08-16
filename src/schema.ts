import { BorderStyle, HeadingLevel } from 'docx';
import { DocxSerializer, MarkSerializer, NodeSerializer } from './serializer';
import { coverColorToHex } from './utils';

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
    state.text((node.text ?? '').replace(/\u200b/g, ''));
  },
  paragraph(state, node) {
    state.renderInline(node);
    state.closeBlock(node);
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
    state.closeBlock(node, { heading });
  },
  blockquote(state, node) {
    state.renderContent(node, {
      style: 'IntenseQuote',
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
    // Kinda hacky, but this works to insert two paragraphs, the first with a break
    state.closeBlock(node, { thematicBreak: true });
    state.closeBlock(node);
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
    state.math(node.attrs.input, { inline: true });
  },
  blocked_latex(state, node) {
    const { id = Date.now(), numbered } = node.attrs;
    state.math(node.attrs.input, { inline: false, numbered, id });
    state.closeBlock(node);
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
      color: mark.attrs.color ? coverColorToHex(mark.attrs.color) : '#000000',
    };
  },
  font_family(state, node, mark) {
    return {
      font: {
        name: mark.attrs.font_family || 'Kaiti SC',
      },
    };
  },
  font_size(state, node, mark) {
    try {
      return {
        size: parseInt(mark.attrs.fontSize || '', 10),
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
    return { subScript: true };
  },
  strike() {
    // doubleStrike!
    return { strike: true };
  },
  underline() {
    return {
      underline: {},
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
