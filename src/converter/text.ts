import { TextRun, XmlComponent, type IRunOptions } from "docx";
import { NodeConverter } from "./node.js";
import type { Node } from "dom-parser";
import type { ConverterContext } from "../context/convertercontext.js";


export class TextConverter extends NodeConverter {
  convert(node: Node, run: IRunOptions, context: ConverterContext): XmlComponent {
    return new TextRun({
      ...run,
      text: node.textContent || "",
    });
  }
}