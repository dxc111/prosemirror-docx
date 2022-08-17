import { AlignmentType } from 'docx';

export default {
  default: {
    heading1: {
      run: {
        font: 'Calibri',
        size: 52,
        bold: true,
        color: '2E3A59',
      },
      paragraph: {
        spacing: { line: 340 },
      },
    },
    heading2: {
      run: {
        font: 'Calibri',
        size: 26,
        bold: true,
        color: '2E3A59',
      },
      paragraph: {
        spacing: { line: 340 },
      },
    },
    heading3: {
      run: {
        font: 'Calibri',
        size: 26,
        bold: true,
        color: '2E3A59',
      },
      paragraph: {
        spacing: { line: 276 },
      },
    },
    heading4: {
      run: {
        font: 'Calibri',
        size: 26,
        bold: true,
        color: '2E3A59',
      },
    },
  },
  paragraphStyles: [
    {
      id: 'NormalPara',
      name: 'Normal Para',
      basedOn: 'Normal',
      next: 'Normal',
      quickFormat: true,
      run: {
        size: 20,
        color: '2E3A59',
      },
    },
    {
      id: 'Aside',
      name: 'Aside',
      basedOn: 'Normal',
      next: 'Normal',
      run: {
        color: '999999',
        italics: true,
        size: 18,
      },
      paragraph: {
        spacing: {
          line: 276,
        },
        alignment: AlignmentType.CENTER,
      },
    },
    {
      id: 'BlockCode',
      name: 'Block Code',
      basedOn: 'Normal',
      next: 'Normal',
      run: {
        color: '6e7d8b',
        italics: true,
        size: 18,
      },
      paragraph: {
        spacing: {
          before: 30,
          after: 30,
          line: 276,
        },
        indent: { left: 250 },
      },
    },
    {
      id: 'IntenseQuote',
      name: 'Intense Quote',
      basedOn: 'Normal',
      next: 'Normal',
      run: {
        italics: true,
      },
      paragraph: {
        indent: { left: 250 },
      },
    },
  ],
};
