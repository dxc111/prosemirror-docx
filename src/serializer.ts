import { Mark, Node as ProsemirrorNode, Schema } from 'prosemirror-model';
import {
  AlignmentType,
  Bookmark,
  Column,
  ColumnBreak,
  CommentRangeEnd,
  CommentRangeStart,
  CommentReference,
  ExternalHyperlink,
  FootnoteReferenceRun,
  HeadingLevel,
  ICommentOptions,
  ImageRun,
  IParagraphOptions,
  IRunOptions,
  ITableCellOptions,
  LineRuleType,
  Math,
  MathRun,
  Paragraph,
  ParagraphChild,
  SectionType,
  SequentialIdentifier,
  Table,
  TableCell,
  TableRow,
  TabStopPosition,
  TabStopType,
  TextRun,
  TextWrappingType,
  WidthType,
} from 'docx';
import { createNumbering, INumbering, NumberingStyles } from './numbering';
import { createDocFromState, createShortId } from './utils';

type Mutable<T> = {
  -readonly [k in keyof T]: T[k];
};

function normalizeText(text: string) {
  return (
    (text || '') // eslint-disable-next-line no-misleading-character-class
      .replace(/[\u200B\u200C\u200D\uFEFF]/g, '')
      // eslint-disable-next-line no-control-regex
      .replace(/[\u0000-\u0008\u000B-\u000C\u000E-\u001F]/g, '')
  );
}

const MAX_IMAGE_WIDTH = 600;
// This is duplicated from @curvenote/schema
export type AlignOptions = 'left' | 'center' | 'right';

export type NodeSerializer<S extends Schema = any> = Record<
  string,
  (
    state: DocxSerializerState<S>,
    node: ProsemirrorNode<S>,
    parent: ProsemirrorNode<S>,
    index: number,
  ) => void
>;

export type MarkSerializer<S extends Schema = any> = Record<
  string,
  (state: DocxSerializerState<S>, node: ProsemirrorNode<S>, mark: Mark<S>) => IRunOptions
>;

interface ImageBuffer {
  arrayBuffer: string | ArrayBuffer;
  width: number;
  height: number;
}

export type Options = {
  getImageBuffer: (src: string) => ImageBuffer;
};

export type IMathOpts = {
  inline?: boolean;
  id?: string | null;
  numbered?: boolean;
};

let currentLink: { link: string; stack: ParagraphChild[] } | undefined;

export class DocxSerializerState<S extends Schema = any> {
  nodes: NodeSerializer<S>;

  options: Options;

  marks: MarkSerializer<S>;

  children: Paragraph[] | any;

  numbering: INumbering[];

  nextRunOpts?: IRunOptions;

  current: ParagraphChild[] | any = [];

  currentBlockNode = '';

  currentLink?: { link: string; children: IRunOptions[] };

  comments: ICommentOptions[] = [];

  pageBreak = 'hr';

  // Optionally add options
  nextParentParagraphOpts?: IParagraphOptions;

  currentNumbering?: { reference: string; level: number };

  private footnoteIdx: number;

  private footnoteState = '';

  private footnoteIds: string[];

  private maxImageWidth = 600;

  public fullCiteContents: Record<string, string>;

  transformHtmlToNode: (html: string) => null;

  constructor(
    nodes: NodeSerializer<S>,
    marks: MarkSerializer<S>,
    options: Options,
    fullCiteContents: Record<string, string>,
    pageBreak = 'hr',
    private numberingStyles: Record<NumberingStyles, any> | null = null,
    private cslFormatService: any = null,
    private bibliographyTitle = 'Bibliography',
    footnoteState = 'disable',
    transformHtmlToNode = (html: string) => null,
  ) {
    this.nodes = nodes;
    this.marks = marks;
    this.options = options ?? {};
    this.children = [];
    this.numbering = [];
    this.footnoteIdx = 0;
    this.footnoteIds = [];
    this.fullCiteContents = fullCiteContents;
    this.pageBreak = pageBreak;
    this.footnoteState = footnoteState;
    this.transformHtmlToNode = transformHtmlToNode;
  }

  renderContent(parent: ProsemirrorNode, opts?: IParagraphOptions) {
    parent.forEach((node, _, i) => {
      if (opts) this.addParagraphOptions(opts);
      this.render(node, parent, i);
    });
  }

  render(node: ProsemirrorNode<S>, parent: ProsemirrorNode<S>, index: number) {
    const target = this.nodes[node.type.name] || this.nodes.default;
    if (!target) throw new Error(`Token type \`${node.type.name}\` not supported by Word renderer`);
    target(this, node, parent, index);
  }

  renderMarks(node: ProsemirrorNode<S>, marks: Mark[]): IRunOptions {
    return marks
      .map((mark) => {
        return this.marks[mark.type.name]?.(this, node, mark);
      })
      .reduce((a, b) => ({ ...a, ...b }), {});
  }

  renderParagraphHtml(html: string) {
    // eslint-disable-next-line no-param-reassign
    html = normalizeText(html);
    const node = this.transformHtmlToNode(html);
    if (node) {
      // this.render(node, node, 0);
      const cache = this.current;
      this.current = [];
      this.renderInline(node);
      const res = this.current;
      this.current = cache;

      return res;
    }
    return [new TextRun(html)];
  }

  openLink(href: string) {
    this.addRunOptions({ style: 'Hyperlink' });
    // TODO: https://github.com/dolanmiu/docx/issues/1119
    // Remove the if statement here and oneLink!
    // const oneLink = true;
    // if (!oneLink) {
    //   closeLink();
    // } else {
    //   if (currentLink && sameLink) return;
    //   if (currentLink && !sameLink) {
    //     // Close previous, and open a new one
    //     closeLink();
    //   }
    // }
    currentLink = {
      link: href,
      stack: this.current,
    };
    this.current = [];
  }

  curIdx = 0;

  wrapComment(node: ProsemirrorNode) {
    if (node.type.name === 'comment') {
      let time = Date.now();
      try {
        time = parseInt(node.attrs.createDate, 10);
      } catch (e) {
        // eslint-disable-next-line no-console
        console.error('Error parsing comment time: ', e);
      }

      this.comments.push({
        id: this.curIdx,
        children: [
          new Paragraph({
            children: [
              new TextRun({
                text: node.attrs.comment,
              }),
            ],
          }),
        ],
        date: new Date(time),
      });
      this.current.push(new CommentRangeStart(this.curIdx));
      this.renderInline(node);
      this.current.push(
        new CommentRangeEnd(this.curIdx),
        new TextRun({
          children: [new CommentReference(this.curIdx)],
        }),
      );
      this.curIdx += 1;
    }
  }

  closeLink() {
    if (!currentLink) return;
    const hyperlink = new ExternalHyperlink({
      link: currentLink.link,
      // child: this.current[0],
      children: this.current,
    });
    this.current = [...currentLink.stack, hyperlink];
    currentLink = undefined;
  }

  openTable() {}

  closeTable() {}

  hierarchy_title(node: ProsemirrorNode<S>) {
    if (this.pageBreak === 'page' && this.children.length > 0) {
      this.addParagraphOptions({ pageBreakBefore: true });
    }
    if (node.content.size > 0) {
      this.renderInline(node);
      const heading = [
        HeadingLevel.HEADING_1,
        HeadingLevel.HEADING_2,
        HeadingLevel.HEADING_3,
        HeadingLevel.HEADING_4,
        HeadingLevel.HEADING_5,
        HeadingLevel.HEADING_6,
      ][node.attrs.level - 1];
      this.closeBlock(node, { heading, style: `heading${node.attrs.level}` });
    }
  }

  horizontal_rule(node: ProsemirrorNode<S>) {
    if (this.pageBreak === 'hr') {
      this.addParagraphOptions({ pageBreakBefore: true });
    } else {
      // Kinda hacky, but this works to insert two paragraphs, the first with a break
      this.closeBlock(node, { thematicBreak: true });
      this.closeBlock(node);
    }
  }

  renderInline(parent: ProsemirrorNode<S>) {
    // Pop the stack over to this object when we encounter a link, and closeLink restores it
    // let currentLink: { link: string; stack: ParagraphChild[] } | undefined;
    // const closeLink = () => {
    //   if (!currentLink) return;
    //   const hyperlink = new ExternalHyperlink({
    //     link: currentLink.link,
    //     // child: this.current[0],
    //     children: this.current,
    //   });
    //   this.current = [...currentLink.stack, hyperlink];
    //   currentLink = undefined;
    // };
    // const openLink = (href: string) => {
    //   const sameLink = href === currentLink?.link;
    //   this.addRunOptions({ style: 'Hyperlink' });
    //   // TODO: https://github.com/dolanmiu/docx/issues/1119
    //   // Remove the if statement here and oneLink!
    //   const oneLink = true;
    //   if (!oneLink) {
    //     closeLink();
    //   } else {
    //     if (currentLink && sameLink) return;
    //     if (currentLink && !sameLink) {
    //       // Close previous, and open a new one
    //       closeLink();
    //     }
    //   }
    //   currentLink = {
    //     link: href,
    //     stack: this.current,
    //   };
    //   this.current = [];
    // };
    const progress = (node: ProsemirrorNode<S>, offset: number, index: number) => {
      // const links: ProsemirrorNode[] = [];
      // node.forEach((child) => {
      //   if (child.type.name === 'link') {
      //     links.push(child);
      //     return false;
      //   }
      //   return true;
      // });
      // const hasLink = links.length > 0;
      // if (hasLink) {
      //   openLink(links[0].attrs.href);
      // } else if (!hasLink && currentLink) {
      //   closeLink();
      // }
      if (node.isText) {
        this.text(node.text, this.renderMarks(node, node.marks));
      } else {
        this.render(node, parent, index);
      }
    };
    parent.forEach(progress);
    // Must call close at the end of everything, just in case
    // closeLink();
  }

  renderList(node: ProsemirrorNode<S>, style: NumberingStyles) {
    if (!this.currentNumbering) {
      const nextId = createShortId();
      this.numbering.push(createNumbering(nextId, style, this.numberingStyles?.[style] || null));
      this.currentNumbering = { reference: nextId, level: 0 };
    } else {
      const { reference, level } = this.currentNumbering;
      this.currentNumbering = { reference, level: level + 1 };
    }
    this.renderContent(node, {
      style: `${style}list`,
    });
    if (this.currentNumbering.level === 0) {
      delete this.currentNumbering;
    } else {
      const { reference, level } = this.currentNumbering;
      this.currentNumbering = { reference, level: level - 1 };
    }
  }

  // This is a pass through to the paragraphs, etc. underneath they will close the block
  renderListItem(node: ProsemirrorNode<S>) {
    if (!this.currentNumbering) throw new Error('Trying to create a list item without a list?');
    this.addParagraphOptions({ numbering: this.currentNumbering });
    // this.renderContent(node);
    let onlyParagraph = true;
    node.forEach((n, _, i) => {
      if (n.type.name === 'paragraph' && onlyParagraph) {
        if (this.current.length > 0) {
          // add a break between paragraphs
          this.current.push(new TextRun({ break: 1 }));
        }
        this.renderInline(n);
      } else {
        if (this.current.length > 0) {
          this.closeBlock(n);
        }
        onlyParagraph = false;
        this.render(n, node, i);
      }
    });
    if (onlyParagraph) {
      this.closeBlock(node);
    }
  }

  addParagraphOptions(opts: IParagraphOptions) {
    this.nextParentParagraphOpts = { ...this.nextParentParagraphOpts, ...opts };
  }

  addRunOptions(opts: IRunOptions) {
    this.nextRunOpts = { ...this.nextRunOpts, ...opts };
  }

  text(text: string | null | undefined, opts?: IRunOptions) {
    const textNormalized = normalizeText(text || '');
    if (!text) return;
    this.current.push(
      new TextRun({
        text: textNormalized,
        ...(currentLink ? { style: 'Hyperlink' } : {}),
        ...this.nextRunOpts,
        ...opts,
      }),
    );
    delete this.nextRunOpts;
  }

  math(latex: string, opts: IMathOpts = { inline: true }) {
    if (opts.inline || !opts.numbered) {
      this.current.push(new Math({ children: [new MathRun(latex)] }));
      return;
    }
    const id = opts.id ?? createShortId();
    this.current = [
      new TextRun('\t'),
      new Math({
        children: [new MathRun(latex)],
      }),
      new TextRun('\t('),
      new Bookmark({
        id,
        children: [new SequentialIdentifier('Equation')],
      }),
      new TextRun(')'),
    ];
    this.addParagraphOptions({
      tabStops: [
        {
          type: TabStopType.CENTER,
          position: TabStopPosition.MAX / 2,
        },
        {
          type: TabStopType.RIGHT,
          position: TabStopPosition.MAX,
        },
      ],
    });
  }

  bib_cite(node: ProsemirrorNode<S>) {
    try {
      if (node.type.name === 'group_bio_citation') {
        this.current.push(new TextRun(node.attrs.cache));
      }
      if (this.cslFormatService) {
        const cite = this.cslFormatService.getCitationByIdSync(
          node.attrs.reference || node.attrs.metadataId,
          'text',
        );
        this.current.push(new TextRun(cite));
      }
    } catch (error) {
      // eslint-disable-next-line no-console
      console.error(error);
    }
  }

  bibliography(node: ProsemirrorNode<S>) {
    try {
      if (node.type.name === 'bibliography') {
        if (this.cslFormatService) {
          const bib = this.cslFormatService?.getBibliographyByIdSync(undefined, 'text', true);

          this.closeBlock(node);
          this.current.push(new TextRun(this.bibliographyTitle));
          this.closeBlock(node, { style: 'BibliographyTitle' });

          bib.forEach(([_, bibliography]: any) => {
            this.current.push(new TextRun(bibliography));
            this.closeBlock(node, { style: 'Bibliography' });
          });
        }
      } else {
        this.closeBlock(node);
        this.current.push(new TextRun(this.bibliographyTitle));
        this.closeBlock(node, { style: 'BibliographyTitle' });

        node.content.forEach((n) => {
          this.current.push(new TextRun(n.textContent));
          this.addParagraphOptions({ style: 'Bibliography' });
          this.closeBlock(node);
        });
      }
    } catch (error) {
      // eslint-disable-next-line no-console
      console.error(error);
    }
  }

  footnoteRef(id: string) {
    if (this.footnoteState === 'disable') return;
    this.footnoteIds.push(id);
    this.footnoteIdx += 1;
    this.current.push(
      this.footnoteState === 'endnotes'
        ? new TextRun({
            text: `${this.footnoteIdx}`,
            style: 'FootnoteReference',
          })
        : new FootnoteReferenceRun(this.footnoteIdx),
    );
  }

  imageWithType() {}

  image(src: string, align: AlignOptions = 'center', widthPercent = 90) {
    const { arrayBuffer, width: rawW, height: rawH } = this.options.getImageBuffer(src);

    const aspect = rawH / rawW;
    const width = this.maxImageWidth * (widthPercent / 100);

    this.current.push(
      new ImageRun({
        data: arrayBuffer,
        transformation: {
          width,
          height: width * aspect,
        },
        // floating: {
        //   horizontalPosition: {
        //     offset: 0,
        //   },
        //   verticalPosition: {
        //     offset: 0,
        //   },
        //   wrap: {
        //     type: TextWrappingType.TOP_AND_BOTTOM,
        //   },
        // },
      }),
    );
    let alignment: AlignmentType;
    switch (align) {
      case 'right':
        alignment = AlignmentType.RIGHT;
        break;
      case 'left':
        alignment = AlignmentType.LEFT;
        break;
      default:
        alignment = AlignmentType.CENTER;
    }
    this.addParagraphOptions({
      alignment,
      style: 'Normal',
    });
  }

  imageInline(src: string, maxHeight = 0) {
    const { arrayBuffer, width, height } = this.options.getImageBuffer(src);

    if (maxHeight) {
      const aspect = height / width;
      const newWidth = maxHeight / aspect;
      this.current.push(
        new ImageRun({
          data: arrayBuffer,
          transformation: {
            width: newWidth,
            height: maxHeight,
          },
        }),
      );
      return;
    }

    this.current.push(
      new ImageRun({
        data: arrayBuffer,
        transformation: {
          width,
          height,
        },
      }),
    );
  }

  addAside(text = '') {
    this.children.push(
      new Paragraph({
        text,
        style: 'Aside',
      }),
    );
  }

  addCodeBlock(node: ProsemirrorNode) {
    if (node.textContent) {
      // this.children.push(new Paragraph({ text: '' }));
      // node.textContent.split('\n').forEach((text) => {
      //   this.children.push(
      //     new Paragraph({
      //       text,
      //       style: 'BlockCode',
      //     }),
      //   );
      // });
      this.children.push(
        new Paragraph({
          children: normalizeText(node.textContent)
            .split('\n')
            .map((text, idx) => new TextRun({ text, break: idx > 0 ? 1 : undefined })),
          style: 'BlockCode',
        }),
      );
      // this.children.push(new Paragraph({ text: '' }));
    }
  }

  captionLabel(id: string, kind: 'Figure' | 'Table') {
    this.current.push(
      new Bookmark({
        id,
        children: [new TextRun(`${kind} `), new SequentialIdentifier(kind)],
      }),
    );
  }

  table(node: ProsemirrorNode<S>) {
    const actualChildren = this.children;
    const rows: TableRow[] = [];
    let percent = 0;
    let colCount = 0;
    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore
    node.content.forEach(({ content: rowContent }) => {
      const cells: TableCell[] = [];
      // Check if all cells are headers in this row
      let tableHeader = true;
      rowContent.forEach((cell: { type: { name: string } }) => {
        if (cell.type.name !== 'table_header') {
          tableHeader = false;
        }
      });
      // This scales images inside of tables
      this.maxImageWidth = MAX_IMAGE_WIDTH / rowContent.childCount;
      percent = percent || 100 / (rowContent.childCount || 1);
      colCount = rowContent.childCount;
      rowContent.forEach((cell: ProsemirrorNode<S>) => {
        this.children = [];
        this.renderContent(cell, { style: 'TableCell' });
        const tableCellOpts: Mutable<ITableCellOptions> = {
          children: this.children,
          width: {
            type: WidthType.PERCENTAGE,
            size: percent,
          },
        };
        const colspan = cell.attrs.colspan ?? 1;
        const rowspan = cell.attrs.rowspan ?? 1;
        if (colspan > 1) tableCellOpts.columnSpan = colspan;
        if (rowspan > 1) tableCellOpts.rowSpan = rowspan;
        cells.push(new TableCell(tableCellOpts));
      });
      rows.push(new TableRow({ children: cells, tableHeader }));
    });
    this.maxImageWidth = MAX_IMAGE_WIDTH;
    const table = new Table({
      rows,
      columnWidths: new Array(colCount).fill(0).map(() => 9010 / (colCount || 1)),
      // width: {
      //   type: WidthType.DXA,
      //   size: 9010,
      // },
    });
    // if (table instanceof Paragraph) {
    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore
    actualChildren.push(table);
    // }
    // If there are multiple tables, this seperates them
    actualChildren.push(new Paragraph(''));
    this.children = actualChildren;
  }

  columns(node: ProsemirrorNode<S>) {
    if (node.childCount < 1) return;
    const actualChildren = this.children;
    const columnsItems: Paragraph[] = [];
    const columnsWidth: Column[] = [];

    node.content.forEach((column: ProsemirrorNode<S>, _, idx) => {
      this.children = [];

      if (idx > 0 && idx < node.childCount - 1) {
        // const lastParagraph = columnsItems[columnsItems.length - 1];
        // if (lastParagraph && lastParagraph instanceof Paragraph) {
        //   lastParagraph.addChildElement(new ColumnBreak());
        // } else {
        //   columnsItems.push(new Paragraph({ children: [new ColumnBreak()] }));
        // }
        columnsItems.push(
          new Paragraph({
            children: [new ColumnBreak()],
            spacing: {
              line: 0,
              lineRule: LineRuleType.EXACT,
              before: 0,
              after: 0,
            },
          }),
        );
      }

      columnsWidth.push(new Column({ width: (parseFloat(column.attrs.basis) / 100) * 9010 }));
      this.maxImageWidth = (MAX_IMAGE_WIDTH * parseFloat(column.attrs.basis)) / 100;
      this.renderContent(column);
      // column.content.forEach((child) => {
      //   this.renderContent(child);
      // });

      columnsItems.push(...this.children);
    });

    actualChildren.push({
      properties: {
        type: SectionType.CONTINUOUS,
        column: {
          space: 708,
          count: node.childCount,
          equalWidth: false,
          children: columnsWidth,
        },
      },
      children: columnsItems,
    });
    actualChildren.push(new Paragraph(''));
    this.children = actualChildren;
    this.maxImageWidth = MAX_IMAGE_WIDTH;
  }

  closeBlock(node: ProsemirrorNode<S>, props?: IParagraphOptions) {
    const paragraph = new Paragraph({
      children: this.current,
      style: 'NormalPara',
      ...this.nextParentParagraphOpts,
      ...props,
    });
    this.current = [];
    delete this.nextParentParagraphOpts;
    this.children.push(paragraph);
  }
}

export class DocxSerializer<S extends Schema = any> {
  nodes: NodeSerializer<S>;

  marks: MarkSerializer<S>;

  constructor(nodes: NodeSerializer<S>, marks: MarkSerializer<S>) {
    this.nodes = nodes;
    this.marks = marks;
  }

  serialize(
    content: ProsemirrorNode<S>,
    options: Options,
    footerText = '',
    footnotes: string[] = [],
    pageOptions: any,
    fullCiteContents: Record<string, string>,
    externalStyles: any = null,
    numberingStyles: Record<NumberingStyles, any> | null = null,
    cslFormatService: any = null,
    bibliographyTitle = 'Bibliography',
    footnoteTitle = 'Footnotes',
    transformHtmlToNode?: (html: string) => any,
    log?: (...args: any[]) => void,
  ) {
    const enableFootnotes = !!pageOptions?.footnotes;
    const isEndNotes =
      enableFootnotes &&
      (pageOptions?.footnotesPosition === 'page' ||
        pageOptions?.footnotesPosition === 'before_bib');
    let footnoteState = enableFootnotes ? 'enable' : 'disable';
    if (isEndNotes) {
      footnoteState = 'endnotes';
    }

    if (log) {
      log('isEndNotes: ', isEndNotes);
    }
    const state = new DocxSerializerState<S>(
      this.nodes,
      this.marks,
      options,
      fullCiteContents,
      pageOptions?.splitPage || 'hr',
      numberingStyles,
      cslFormatService,
      bibliographyTitle,
      footnoteState,
      transformHtmlToNode,
    );
    // eslint-disable-next-line no-param-reassign
    footnotes = footnotes.map((f) => state.renderParagraphHtml(f));
    state.renderContent(content);
    const f: Record<number, any> = footnotes.reduce((acc: Record<number, any>, cur, idx) => {
      acc[idx + 1] = {
        children: [
          new Paragraph({
            style: 'FootnoteList',
            children: (Array.isArray(cur) ? cur : [cur]) as any,
          }),
        ],
      };
      return acc;
    }, {});

    try {
      if (isEndNotes && footnotes.length > 0) {
        const idx = state.children.findIndex((c: any) => c === '[[THIS_IS_A_FOOTNOTES_HOLE]]');
        if (idx > -1 && pageOptions?.footnotesPosition === 'before_bib') {
          if (log) {
            log('Bibliography is Last');
          }
          state.children.splice(
            idx,
            1,
            new Paragraph({ children: [] }),
            new Paragraph({
              text: footnoteTitle,
              style: 'BibliographyTitle',
            }),
            ...footnotes.map(
              (footnote, i) =>
                new Paragraph({
                  style: 'Bibliography',
                  children: [
                    new TextRun(`${i + 1}. `),
                    ...(Array.isArray(footnote) ? footnote : [footnote]),
                  ],
                }),
            ),
          );
        } else {
          state.children.push(
            new Paragraph({ children: [] }),
            new Paragraph({
              text: footnoteTitle,
              style: 'BibliographyTitle',
            }),
            ...footnotes.map(
              (footnote, i) =>
                new Paragraph({
                  style: 'Bibliography',
                  children: [
                    new TextRun(`${i + 1}. `),
                    ...(Array.isArray(footnote) ? footnote : [footnote]),
                  ],
                }),
            ),
          );
        }
      }
    } catch (e) {
      if (log) {
        log('Error adding footnotes: ', e);
      }
    } finally {
      const index = state.children.findIndex((c: any) => c === '[[THIS_IS_A_FOOTNOTES_HOLE]]');
      if (index > -1) {
        state.children.splice(index, 1);
      }
    }

    return createDocFromState(
      state,
      footerText,
      isEndNotes || !enableFootnotes ? {} : f,
      pageOptions,
      options.getImageBuffer,
      externalStyles,
    );
  }
}
