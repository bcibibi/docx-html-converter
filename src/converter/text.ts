import { TextRun, XmlComponent, type IRunOptions } from "docx";
import { NodeConverter } from "./node";
import type { Node } from "dom-parser";
import type { ConverterContext } from "../context/convertercontext";


export class TextConverter extends NodeConverter {
  convert(node: Node, run: IRunOptions, children: XmlComponent[], context: ConverterContext): XmlComponent {
    return new TextRun({
      ...run,
      text: node.textContent || "",
    });
  }
}