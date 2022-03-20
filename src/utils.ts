import {
  AlignmentType,
  Document,
  Footer,
  INumberingOptions,
  ISectionOptions,
  Packer,
  Paragraph,
  SectionType,
} from 'docx';
import { Node as ProsemirrorNode } from 'prosemirror-model';

export function createShortId() {
  return Math.random().toString(36).substr(2, 9);
}

export function createDocFromState(
  state: {
    numbering: INumberingOptions['config'];
    children: ISectionOptions['children'];
  },
  footerText?: string,
) {
  return new Document({
    numbering: {
      config: state.numbering,
    },
    sections: [
      {
        properties: {
          type: SectionType.CONTINUOUS,
        },
        children: state.children,
        footers: footerText
          ? {
              default: new Footer({
                children: [new Paragraph({ text: footerText, alignment: AlignmentType.CENTER })],
              }),
            }
          : undefined,
      },
    ],
  });
}

export function writeDocx(doc: Document, write: (buffer: Blob) => void) {
  Packer.toBlob(doc).then(write);
}

export function getLatexFromNode(node: ProsemirrorNode): string {
  let math = '';
  node.forEach((child) => {
    if (child.isText) math += child.text;
    // TODO: improve this as we may have other things in the future
  });
  return math;
}

export function coverColorToHex(color: string) {
  try {
    const el = document.createElement('div');
    el.style.display = 'none';
    el.style.position = 'fixed';
    el.style.color = color;
    document.body.appendChild(el);
    const rgb = window.getComputedStyle(el).color.replace(/rgba?/, '');
    document.body.removeChild(el);

    return `#${rgb
      .slice(1, -1)
      .split(',')
      .slice(0, 3)
      .map((c) => (+c).toString(16).padStart(2, '0'))
      .join('')}`;
  } catch (e) {
    return color;
  }
}
