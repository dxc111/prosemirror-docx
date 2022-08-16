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
  footnotes: Record<number, any> = {},
) {
  console.log('createDocFromState', state);

  const sections = state.children.reduce((res: any[], cur: any) => {
    if (!cur.properties?.column) {
      if (res[res.length - 1] && !res[res.length - 1].properties?.column) {
        res[res.length - 1].children.push(cur);
      } else {
        res.push({
          properties: {
            type: SectionType.CONTINUOUS,
          },
          children: [cur],
          footers: footerText
            ? {
                default: new Footer({
                  children: [new Paragraph({ text: footerText, alignment: AlignmentType.CENTER })],
                }),
              }
            : undefined,
        });
      }
    } else {
      // eslint-disable-next-line no-param-reassign
      cur.footers = footerText
        ? {
            default: new Footer({
              children: [new Paragraph({ text: footerText, alignment: AlignmentType.CENTER })],
            }),
          }
        : undefined;
      res.push(cur);
    }
    return res;
  }, []);
  console.log(sections);
  return new Document({
    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    footnotes, // @ts-ignore
    comments: { children: state.comments },
    styles: {
      paragraphStyles: [
        {
          id: 'aside',
          name: 'Aside',
          basedOn: 'Normal',
          next: 'Normal',
          run: {
            color: '999999',
            italics: true,
            size: 14,
          },
          paragraph: {
            spacing: {
              line: 276,
            },
            alignment: AlignmentType.CENTER,
          },
        },
      ],
    },
    numbering: {
      config: state.numbering,
      // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    }, // @ts-ignore
    sections,
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
