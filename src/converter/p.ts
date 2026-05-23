import { Paragraph, XmlComponent, type IRunOptions } from "docx";
import { NodeConverter, type GetChildrenFct } from "./node";
import type { Node } from "dom-parser";
import type { ConverterContext } from "../context/convertercontext";

export class PConverter extends NodeConverter {

  convert(node: Node, run: IRunOptions, context: ConverterContext, children: GetChildrenFct): XmlComponent {
    return new Paragraph({
      children: children(node, run, context),
      run
    });
  }
}