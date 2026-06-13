import type { Node } from "dom-parser";
import { NodeConverter, type GetChildrenFct } from "./node";
import { TextRun, type IRunOptions, type XmlComponent } from "docx";
import type { ConverterContext } from "../context/convertercontext";


export class IConverter extends NodeConverter {
  async convert(node: Node, run: IRunOptions, context: ConverterContext, children: GetChildrenFct): Promise<XmlComponent> {
    run = {...run, italics: true};
    return new TextRun({
      ...run,
      children: await children(node, run, context)
    });
  }
}