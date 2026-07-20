import type { IRunOptions, XmlComponent } from "docx";
import type { Node } from "dom-parser";
import { NodeConverter, type GetChildrenFct } from "./node.js";
import type { ConverterContext } from "../context/convertercontext.js";

export class OLConverter extends NodeConverter {

  convert(node: Node, run: IRunOptions, context: ConverterContext, children: GetChildrenFct): Promise<XmlComponent[]> {
    return children(node, run, context);
  }

}