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
  .fill(['●', '○', '■'])
  // .fill(['●', '●', '●'])
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

export function createNumbering(reference: string, style: NumberingStyles): INumbering {
  return {
    reference,
    levels: styles[style],
  };
}
