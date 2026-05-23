import { Paragraph, type IRunOptions, type XmlComponent } from "docx";
import type { Node } from "dom-parser";
import type { ConverterContext } from "../context/convertercontext";
import { NodeConverter, type GetChildrenFct } from "./node";
import debug from "debug";

const log = debug("docxhtml:li");

export class LIConverter extends NodeConverter {

  convert(node: Node, run: IRunOptions, context: ConverterContext, children: GetChildrenFct): XmlComponent {
    const isBullet = this.isBullet(node);
    const isNumbered = this.isNumbered(node);
    log("isBullet", isBullet, "isNumbered", isNumbered);
    return new Paragraph({
      children: children(node, run, context),
      run,
      ...(isBullet ? {
        bullet: {
          level: this.level(node),
        },
      } : isNumbered ? {
         numbering: {
          reference: context.numbering,
          level: this.level(node),
        },
      } : {}),
    });
  }

  isBullet(node: Node): boolean {
    const parent = node.parentNode;
    if (parent) {
      if (parent.nodeName.toUpperCase() === "UL") {
        return true;
      }
      if (parent.nodeName.toUpperCase() === "OL") {
        return false;
      }
      return this.isBullet(parent);
    }
    return false;
  }

  isNumbered(node: Node): boolean {
    const parent = node.parentNode;
    if (parent) {
      if (parent.nodeName.toUpperCase() === "UL") {
        return false;
      }
      if (parent.nodeName.toUpperCase() === "OL") {
        return true;
      }
      return this.isNumbered(parent);
    }
    return false;
  }

  level(node: Node, level = -1): number {
    const parent = node.parentNode;
    if (parent) {
      log("parent", parent.nodeName, "level", level);
      if (parent.nodeName.toUpperCase() === "OL") {
        level = level + 1;
      }
      if (parent.nodeName.toUpperCase() === "UL") {
        level = level + 1;
      }
      return this.level(parent, level);
    }
    return level;
  }
  
}