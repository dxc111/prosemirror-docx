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
  VerticalAlign,
} from 'docx';
import { Node as ProsemirrorNode } from 'prosemirror-model';
import styles from './styles';

type DocxImageType = 'jpg' | 'png' | 'gif' | 'bmp';

export function getImageType(src: string, data?: unknown): DocxImageType {
  let bytes: Uint8Array | undefined;
  if (
    data instanceof ArrayBuffer ||
    Object.prototype.toString.call(data) === '[object ArrayBuffer]'
  ) {
    bytes = new Uint8Array(data as ArrayBuffer);
  } else if (ArrayBuffer.isView(data)) {
    bytes = new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  }

  if (bytes) {
    if (bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return 'jpg';
    if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) {
      return 'png';
    }
    if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46 && bytes[3] === 0x38) {
      return 'gif';
    }
    if (bytes[0] === 0x42 && bytes[1] === 0x4d) return 'bmp';
  }

  const value = String(src || '').toLowerCase();
  const dataType = value.match(/^data:image\/(png|jpe?g|gif|bmp)/)?.[1];
  const extension = dataType || value.match(/\.([a-z0-9]+)(?:[?#]|$)/)?.[1];

  if (extension === 'jpeg' || extension === 'jpg') return 'jpg';
  if (extension === 'gif' || extension === 'bmp' || extension === 'png') return extension;
  return 'png';
}

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
  externalStyles: any = null,
) {
  // eslint-disable-next-line @typescript-eslint/no-use-before-define
  const pageMargin = getPageMargin(pageOptions?.margin || null);
  // 对多栏的支持
  const sections = state.children.reduce((res: any[], cur: any) => {
    if (!cur.properties?.column) {
      if (res[res.length - 1] && !res[res.length - 1].properties?.column) {
        res[res.length - 1].children.push(cur);
      } else {
        res.push({
          properties: {
            type: SectionType.CONTINUOUS,
            page: pageMargin || {},
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
  pageSection.properties = { page: pageMargin || {} };
  sections.unshift(pageSection);

  return new Document({
    background: undefined,
    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    footnotes, // @ts-ignore
    comments: { children: state.comments },
    styles: externalStyles || styles,
    numbering: {
      config: state.numbering,
      // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    }, // @ts-ignore
    sections,
    // externalStyles: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    // <w:styles xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    //     xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    //     xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    //     xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    //     xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w14 w15">
    //     <w:style w:type="paragraph" w:customStyle="1" w:styleId="BlockCode">
    //         <w:name w:val="BlockCode" />
    //         <w:basedOn w:val="Normal" />
    //         <w:qFormat />
    //         <w:rsid w:val="00BE7EA6" />
    //         <w:pPr>
    //             <w:widowControl w:val="0" />
    //             <w:shd w:val="clear" w:color="auto" w:fill="EDEDED" w:themeFill="accent3"
    //                 w:themeFillTint="33" />
    //             <w:spacing w:before="120" w:after="120" />
    //             <w:ind w:leftChars="240" w:rightChars="240" />
    //         </w:pPr>
    //         <w:rPr>
    //             <w:rFonts w:ascii="Menlo" w:eastAsia="Menlo" w:hAnsi="Menlo" />
    //             <w:color w:val="7B7B7B" w:themeColor="accent3" w:themeShade="BF" />
    //             <w:sz w:val="18" />
    //         </w:rPr>
    //     </w:style>
    // </w:styles>
    // `,
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

function getPageMargin(margin: any) {
  if (margin) {
    return {
      margin: {
        top: convertMillimetersToTwip(margin.top * 10),
        header: 50,
        bottom: convertMillimetersToTwip(margin.bottom * 10),
        left: convertMillimetersToTwip(margin.left * 10),
        right: convertMillimetersToTwip(margin.right * 10),
      },
    };
  }
  return null;
}

function getHeaderAndFooter(pageOptions: any = {}, getImageBuffer: any) {
  const section: any = { children: [] };

  if (pageOptions.header && pageOptions.header.isActive) {
    let image = null;
    if (pageOptions.header.image) {
      const { arrayBuffer, width: rawW, height: rawH } = getImageBuffer(pageOptions.header.image);

      const aspect = rawH / rawW;
      const height = 30;

      image = new ImageRun({
        type: getImageType(pageOptions.header.image, arrayBuffer),
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
            style: 'Header2',
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
            style: 'Footer2',
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
