import {
  AlignmentType,
  convertMillimetersToTwip,
  Document,
  Footer,
  Header,
  ImageRun,
  INumberingOptions,
  ISectionOptions,
  Packer,
  PageNumber,
  Paragraph,
  SectionType,
  TextRun,
} from 'docx';
import { Node as ProsemirrorNode } from 'prosemirror-model';
import styles from './styles';

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
  pageOptions: any = null,
  getImageBuffer: any = async () => null,
) {
  // 对多栏的支持
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
        });
      }
    } else {
      res.push(cur);
    }
    return res;
  }, []);

  // eslint-disable-next-line @typescript-eslint/no-use-before-define
  const pageSection = getHeaderAndFooter(pageOptions, getImageBuffer);

  sections.unshift(pageSection);

  return new Document({
    background: undefined,
    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    footnotes, // @ts-ignore
    comments: { children: state.comments },
    styles,
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

function getHeaderAndFooter(pageOptions: any = {}, getImageBuffer: any) {
  const section: any = { children: [] };
  if (pageOptions.margin) {
    section.properties = {
      page: {
        margin: {
          top: convertMillimetersToTwip(pageOptions.margin.top * 10),
          bottom: convertMillimetersToTwip(pageOptions.margin.bottom * 10),
          left: convertMillimetersToTwip(pageOptions.margin.left * 10),
          right: convertMillimetersToTwip(pageOptions.margin.right * 10),
        },
      },
    };
  }

  if (pageOptions.header && pageOptions.header.isActive) {
    let image = null;
    if (pageOptions.header.image) {
      const { arrayBuffer, width: rawW, height: rawH } = getImageBuffer(pageOptions.header.image);

      const aspect = rawH / rawW;
      const height = convertMillimetersToTwip(10);
      console.log(height);
      image = new ImageRun({
        data: arrayBuffer,
        transformation: {
          width: height / aspect,
          height,
        },
      });
    }
    section.headers = {
      default: new Header({
        children: [
          new Paragraph({
            children: [
              ...(image ? [image, new TextRun('  ')] : []),
              new TextRun(pageOptions.header.text),
            ],
            // eslint-disable-next-line @typescript-eslint/no-use-before-define
            alignment: getAlignment(pageOptions.header.position),
          }),
        ],
      }),
    };
  }

  if (pageOptions.footer && pageOptions.footer.isActive) {
    section.footers = {
      default: new Footer({
        children: [
          new Paragraph({
            children: [
              new TextRun({
                children: [PageNumber.CURRENT],
              }),
            ],
            // eslint-disable-next-line @typescript-eslint/no-use-before-define
            alignment: getAlignment(pageOptions.footer.position),
          }),
        ],
      }),
    };
  }

  return section;
}

function getAlignment(alignment = '') {
  switch (alignment) {
    case 'right':
      return AlignmentType.RIGHT;
    case 'center':
      return AlignmentType.CENTER;
    default:
      return AlignmentType.LEFT;
  }
}
