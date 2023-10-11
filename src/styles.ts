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
      paragraph: {
        spacing: {
          before: 10,
        },
      },
    },
    {
      id: 'Footer2',
      name: 'Footer2',
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
      name: 'BlockCode',
      basedOn: 'Normal',
      quickFormat: true,
      run: {
        font: 'Menlo',
        color: '282828',
        size: 18,
      },
      paragraph: {
        alignment: AlignmentType.LEFT,
        // eslint-disable-next-line @typescript-eslint/ban-ts-comment
        // @ts-ignore
        border: {
          left: {
            color: 'E6E6E6',
            space: 1,
            style: 'single',
            size: 6,
          },
          right: {
            color: 'E6E6E6',
            space: 1,
            style: 'single',
            size: 6,
          },
          top: {
            color: 'E6E6E6',
            space: 1,
            style: 'single',
            size: 6,
          },
          bottom: {
            color: 'E6E6E6',
            space: 1,
            style: 'single',
            size: 6,
          },
        },
        spacing: {
          before: 276,
          after: 276,
          line: 276,
        },
        indent: {
          left: 240,
          right: 240,
        },
        // @ts-ignore
        shading: {
          fill: 'FCFCFC',
          color: 'auto',
          val: 'clear',
        },
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
