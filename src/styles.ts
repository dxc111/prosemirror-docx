import { AlignmentType } from 'docx';

export default {
  default: {
    heading1: {
      run: {
        font: 'Calibri',
        size: 56,
        bold: true,
        color: '2E3A59',
      },
    },
    heading2: {
      run: {
        font: 'Calibri',
        size: 48,
        bold: true,
        color: '2E3A59',
      },
    },
    heading3: {
      run: {
        font: 'Calibri',
        size: 40,
        bold: true,
        color: '2E3A59',
      },
    },
    heading4: {
      run: {
        font: 'Calibri',
        size: 32,
        bold: true,
        color: '2E3A59',
      },
    },
    heading5: {
      run: {
        font: 'Calibri',
        size: 30,
        bold: true,
        color: '2E3A59',
      },
    },
    heading6: {
      run: {
        font: 'Calibri',
        size: 28,
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
      id: 'Header2',
      name: 'Header2',
      basedOn: 'Normal',
      next: 'Normal',
      run: {
        size: 18,
        color: '58617A',
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
