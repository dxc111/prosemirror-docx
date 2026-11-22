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

const NumberedListTypes: Record<string, (typeof LevelFormat)[keyof typeof LevelFormat]> = {
  decimal: LevelFormat.DECIMAL,
  'lower-alpha': LevelFormat.LOWER_LETTER,
  'upper-alpha': LevelFormat.UPPER_LETTER,
  'lower-roman': LevelFormat.LOWER_ROMAN,
  'upper-roman': LevelFormat.UPPER_ROMAN,
  'simp-chinese-informal': LevelFormat.CHINESE_COUNTING,
};

const BulletListTypes: Record<string, string> = {
  disc: '●',
  circle: '○',
  square: '■',
};

type ListLevel = { type: string; format?: 'dot' | 'paren' | 'parens' | 'comma' | 'fullParens' };

export function getMixedLevels(extraStyles: any): ListLevel[] | null {
  const levels = extraStyles?.levels;
  return Array.isArray(levels) && levels.length ? levels : null;
}

// Keep in sync with listMarkerAffixes in the desktop app's template.ts
function mixedLevelText(format: ListLevel['format'], number: number) {
  if (format === 'parens') return `(%${number})`;
  if (format === 'fullParens') return `（%${number}）`;
  if (format === 'comma') return `%${number}、`;
  return format === 'paren' ? `%${number})` : `%${number}.`;
}

// levelOffset: same-type nesting depth of the list that owns level 0 of this definition
function makeLevels(type: NumberingStyles, extraStyles: any, levelOffset = 0) {
  const mixedLevels = getMixedLevels(extraStyles);
  const mixedLevel = (level: number) =>
    mixedLevels ? mixedLevels[(level + levelOffset) % mixedLevels.length] : null;

  if (type === 'numbered') {
    const listStyleType = styles
      ? NumberedListTypes[extraStyles.listStyleType] || LevelFormat.DECIMAL
      : LevelFormat.DECIMAL;

    return new Array(9).fill(listStyleType).map((format, level) => ({
      level,
      format: NumberedListTypes[mixedLevel(level)?.type || ''] || format,
      text: mixedLevels
        ? mixedLevelText(mixedLevel(level)?.format, level + 1)
        : `${new Array(level + 1).fill(1).reduce((res, _, idx) => `${res}%${idx + 1}.`, '')}`,
      alignment: AlignmentType.START,
      style: {
        run: {
          size: extraStyles?.size || 18,
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

  return new Array(9).fill(LevelFormat.BULLET).map((format, level) => ({
    level,
    format,
    text: BulletListTypes[mixedLevel(level)?.type || extraStyles?.listStyleType] || '●',
    alignment: AlignmentType.START,
    style: {
      run: {
        size: '18',
        color: extraStyles?.listStyleColor || 'auto',
      },
      paragraph: {
        indent: {
          left: convertInchesToTwip(0.25 * (level + 1)),
          hanging: convertInchesToTwip(0.25),
        },
      },
    },
  }));
}

export function createNumbering(
  reference: string,
  style: NumberingStyles,
  extraStyles: any = null,
  levelOffset = 0,
): INumbering {
  let numbering: any = styles?.[style];
  if (extraStyles) {
    const levels = makeLevels(style, extraStyles, levelOffset);
    if (levels) numbering = levels;
  }
  return {
    reference,
    levels: numbering,
  };
}
