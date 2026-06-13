import type { IRunOptions } from "docx";
import type { Node } from "dom-parser";
import cssUnitConverter, { type CSSUnits } from "css-unit-converter";
import { parseColor } from "./color";
import type { ConverterContext } from "../context/convertercontext";

export namespace CSSParser {

  type CSSParserResultFct = (value: string) => IRunOptions;
  type CSSParserResult = CSSParserResultFct | IRunOptions;

  const parsers: Record<string, CSSParserResult> = {
    "font-weight": (value: string) => {
      if (value === 'bold') {
        return {
          bold: true
        }
      }
      return {};
    },
    "font-style": (value: string) => {
      if (value === 'italic') {
        return {
          italics: true
        }
      }
      return {};
    },
    "text-decoration": (value: string) => {
      if (value === 'underline') {
        return {
          underline: {
            type: "single",
            color: "auto"
          }
        }
      }
      if (value === 'line-through') {
        return {
          strike: true
        }
      }
      return {};
    },
    "vertical-align": (value: string) => {
      if (value === 'sub') {
        return {
          subScript: true
        }
      }
      if (value === 'super') {
        return {
          superScript: true
        }
      }
      return {};
    },
    "font-size": (value: string) => {
      const regex = /^(\d+)(\D+)$/;
      const matches = value.match(regex);
      if (matches) {
        const v = parseInt(matches[1] || '');
        const u = matches[2];
        if (!isNaN(v)) {
          const sizePt = cssUnitConverter(v, u as CSSUnits, 'pt');
          return {
            size: sizePt * 2
          }
        }
      }
      return {};
    },
    "color": (value: string) => {
      const color = parseColor(value);
      if (color) {
        return {
          color
        }
      }
      return {};
    }
  };

  export function parse(node: Node, context: ConverterContext): IRunOptions {
    let result: IRunOptions = {};
    const items = context.getCss(node);
    for (const [key, value] of Object.entries(items)) {
      if (key && value) {
        const parser = parsers[key];
        if (parser) {
          if (typeof parser === 'function') {
            result = {
              ...result,
              ...parser(value)
            }
          } else {
            result = {
              ...result,
              ...parser
            }
          }
        }
      }
    }
    return result;
  }
  
}