import type { IRunOptions, XmlComponent } from "docx";
import type { Node } from "dom-parser";
import type { ConverterContext } from "../context/convertercontext";

export abstract class NodeConverter {
  
  abstract convert(node: Node, run: IRunOptions, children: XmlComponent[], context: ConverterContext): XmlComponent | XmlComponent[];

}