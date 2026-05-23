import { TextRun, type IRunOptions, type XmlComponent } from "docx";
import type { Node } from "dom-parser";
import { NodeConverter } from "./node";
import type { ConverterContext } from "../context/convertercontext";

export class SPANConverter extends NodeConverter {

  convert(node: Node, run: IRunOptions, children: XmlComponent[], context: ConverterContext): XmlComponent {
    return new TextRun({
      ...run,
      children,
    });
  }
}