import {
  AlignmentType,
  convertInchesToTwip,
  ILevelsOptions,
  INumberingOptions,
  LevelFormat,
} from 'docx';

export type INumbering = INumberingOptions['config'][0];

function basicIndentStyle(indent: number): Pick<ILevelsOptions, 'style' | 'alignment'> {
  return {
    alignment: AlignmentType.START,
    style: {
      paragraph: {
        indent: { left: convertInchesToTwip(indent), hanging: convertInchesToTwip(0.18) },
      },
    },
  };
}

const numbered = Array(3)
  .fill([LevelFormat.DECIMAL, LevelFormat.DECIMAL, LevelFormat.DECIMAL])
  .flat()
  .map((format, level) => ({
    level,
    format,
    text: `${new Array(level + 1).fill(1).reduce((res, _, idx) => `${res}%${idx + 1}.`, '')}`,
    ...basicIndentStyle((level + 1) / 2),
  }));

const bullets = Array(3)
  // .fill(['●', '○', '■'])
  .fill(['●', '●', '●'])
  .flat()
  .map((text, level) => ({
    level,
    format: LevelFormat.BULLET,
    text,
    ...basicIndentStyle((level + 1) / 2),
  }));

const styles = {
  numbered,
  bullets,
};

export type NumberingStyles = keyof typeof styles;

const NumberedListTypes: Record<string, LevelFormat> = {
  decimal: LevelFormat.DECIMAL,
  'lower-alpha': LevelFormat.LOWER_LETTER,
  'lower-roman': LevelFormat.LOWER_ROMAN,
  'upper-roman': LevelFormat.UPPER_ROMAN,
};

const BulletListTypes: Record<string, string> = {
  disc: '●',
  circle: '○',
  square: '■',
};

function makeLevels(type: NumberingStyles, extraStyles: any) {
  if (type === 'numbered') {
    const listStyleType = styles
      ? NumberedListTypes[extraStyles.listStyleType] || LevelFormat.DECIMAL
      : LevelFormat.DECIMAL;

    return new Array(6).fill(listStyleType).map((format, level) => ({
      level,
      format,
      text: `${new Array(level + 1).fill(1).reduce((res, _, idx) => `${res}%${idx + 1}.`, '')}`,
      alignment: AlignmentType.START,
      style: {
        run: {
          size: '18',
          color: extraStyles?.listStyleColor || 'auto',
        },
        paragraph: {
          indent: {
            left: convertInchesToTwip((level + 1) / 2),
            hanging: convertInchesToTwip(0.18),
          },
        },
      },
    }));
  }

  return new Array(6).fill(LevelFormat.BULLET).map((format, level) => ({
    level,
    format,
    text: BulletListTypes[extraStyles?.listStyleType] || '●',
    alignment: AlignmentType.START,
    style: {
      run: {
        size: '18',
        color: extraStyles?.listStyleColor || 'auto',
      },
      paragraph: {
        indent: {
          left: convertInchesToTwip((level + 1) / 2),
          hanging: convertInchesToTwip(0.18),
        },
      },
    },
  }));
}

export function createNumbering(
  reference: string,
  style: NumberingStyles,
  extraStyles: any = null,
): INumbering {
  let numbering: any = styles?.[style];
  if (extraStyles) {
    const levels = makeLevels(style, extraStyles);
    if (levels) numbering = levels;
  }
  return {
    reference,
    levels: numbering,
  };
}
