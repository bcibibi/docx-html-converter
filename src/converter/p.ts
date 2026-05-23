import { Paragraph, XmlComponent, type IRunOptions } from "docx";
import { NodeConverter } from "./node";
import type { Node } from "dom-parser";
import type { ConverterContext } from "../context/convertercontext";

export class PConverter extends NodeConverter {

  convert(node: Node, run: IRunOptions, children: XmlComponent[], context: ConverterContext): XmlComponent {
    return new Paragraph({
      children,
      run
    });
  }
}