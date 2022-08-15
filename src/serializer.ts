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
  ICommentOptions,
  ImageRun,
  IParagraphOptions,
  IRunOptions,
  ITableCellOptions,
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
  WidthType,
} from 'docx';
import { createNumbering, INumbering, NumberingStyles } from './numbering';
import { createDocFromState, createShortId } from './utils';

type Mutable<T> = {
  -readonly [k in keyof T]: T[k];
};

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

  currentLink?: { link: string; children: IRunOptions[] };

  comments: ICommentOptions[] = [];

  // Optionally add options
  nextParentParagraphOpts?: IParagraphOptions;

  currentNumbering?: { reference: string; level: number };

  private footnoteIdx: number;

  private footnoteIds: string[];

  private maxImageWidth = 600;

  constructor(nodes: NodeSerializer<S>, marks: MarkSerializer<S>, options: Options) {
    this.nodes = nodes;
    this.marks = marks;
    this.options = options ?? {};
    this.children = [];
    this.numbering = [];
    this.footnoteIdx = 0;
    this.footnoteIds = [];
  }

  renderContent(parent: ProsemirrorNode<S>) {
    parent.forEach((node, _, i) => this.render(node, parent, i));
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
      this.comments.push({
        id: this.curIdx,
        text: node.attrs.comment,
        date: new Date(node.attrs.createDate),
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
      this.numbering.push(createNumbering(nextId, style));
      this.currentNumbering = { reference: nextId, level: 0 };
    } else {
      const { reference, level } = this.currentNumbering;
      this.currentNumbering = { reference, level: level + 1 };
    }
    this.renderContent(node);
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
    this.renderContent(node);
  }

  addParagraphOptions(opts: IParagraphOptions) {
    this.nextParentParagraphOpts = { ...this.nextParentParagraphOpts, ...opts };
  }

  addRunOptions(opts: IRunOptions) {
    this.nextRunOpts = { ...this.nextRunOpts, ...opts };
  }

  text(text: string | null | undefined, opts?: IRunOptions) {
    if (!text) return;
    this.current.push(
      new TextRun({
        text,
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

  footnoteRef(id: string) {
    this.footnoteIds.push(id);
    this.footnoteIdx += 1;
    this.current.push(new FootnoteReferenceRun(this.footnoteIdx));
  }

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
    });
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
        this.renderContent(cell);
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
      if (idx !== 0) {
        columnsItems.push(new Paragraph({ children: [new ColumnBreak()] }));
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
  ) {
    const state = new DocxSerializerState<S>(this.nodes, this.marks, options);
    state.renderContent(content);
    const f: Record<number, any> = footnotes.reduce((acc: Record<number, any>, cur, idx) => {
      acc[idx + 1] = { children: [new Paragraph({ children: [new TextRun(cur)] })] };
      return acc;
    }, {});

    return createDocFromState(state, footerText, f);
  }
}
