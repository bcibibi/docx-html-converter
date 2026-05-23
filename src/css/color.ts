import { color } from '@csstools/css-color-parser';
import { parseComponentValue } from '@csstools/css-parser-algorithms';
import { serializeRGB } from '@csstools/css-color-parser';
import { tokenize } from '@csstools/css-tokenizer';
import colorconvert from 'color-convert';
import debug from "debug";

const log = debug("docxhtml:color");

export function parseColor(value: string): string | undefined {
  log(`Parsing color value: ${value}`);
  const hwbComponentValue = parseComponentValue(tokenize({ css: value }));
  if (hwbComponentValue) {
    log (`Parsed HWB component value: ${JSON.stringify(hwbComponentValue)}`);
    const colorData = color(hwbComponentValue);
    if (colorData) {
      log(`Parsed color data: ${JSON.stringify(colorData)}`);
      const rgbComponentValue = serializeRGB(colorData);
      log(`Serialized RGB component value: ${rgbComponentValue}`);
      const match = rgbComponentValue.toString().match(/rgb\((\d+), (\d+), (\d+)\)/);
      log(`Extracted RGB values: ${match}`);
      if (match) {
        const [_, r, g, b] = match.map(Number);
        log(`Parsed RGB values: r=${r}, g=${g}, b=${b}`);
        if (r != undefined && g != undefined && b != undefined) {
          const hex = colorconvert.rgb.hex(r, g, b);
          log(`Converted HEX value: ${hex}`);
          return hex;
        }
      }
    }
  }
}
