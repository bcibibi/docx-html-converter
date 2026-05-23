import type { Node } from "dom-parser";
import { NodeConverter, type GetChildrenFct } from "./node";
import { TextRun, type IRunOptions, type XmlComponent } from "docx";
import type { ConverterContext } from "../context/convertercontext";


export class SUBConverter extends NodeConverter {
  convert(node: Node, run: IRunOptions, context: ConverterContext, children: GetChildrenFct): XmlComponent {
    run = {...run, subScript: true};
    return new TextRun({
      ...run,
      children: children(node, run, context)
    });
  }
}