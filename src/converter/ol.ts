import type { IRunOptions, XmlComponent } from "docx";
import type { Node } from "dom-parser";
import { NodeConverter, type GetChildrenFct } from "./node";
import type { ConverterContext } from "../context/convertercontext";

export class OLConverter extends NodeConverter {

  convert(node: Node, run: IRunOptions, context: ConverterContext, children: GetChildrenFct): Promise<XmlComponent[]> {
    return children(node, run, context);
  }

}