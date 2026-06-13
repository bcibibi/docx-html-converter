import type { IRunOptions, XmlComponent } from "docx";
import type { Node } from "dom-parser";
import type { ConverterContext } from "../context/convertercontext";

export type GetChildrenFct = (node: Node, run: IRunOptions, context: ConverterContext) => Promise<XmlComponent[]>;

export abstract class NodeConverter {
  
  abstract convert(node: Node, run: IRunOptions, context: ConverterContext, children: GetChildrenFct): XmlComponent | XmlComponent[] | Promise<XmlComponent | XmlComponent[]>;

}