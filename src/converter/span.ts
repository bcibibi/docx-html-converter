import { TextRun, type IRunOptions, type XmlComponent } from "docx";
import type { Node } from "dom-parser";
import { NodeConverter, type GetChildrenFct } from "./node.js";
import type { ConverterContext } from "../context/convertercontext.js";

export class SPANConverter extends NodeConverter {

  async convert(node: Node, run: IRunOptions, context: ConverterContext, children: GetChildrenFct): Promise<XmlComponent> {
    return new TextRun({
      ...run,
      children: await children(node, run, context),
    });
  }
}