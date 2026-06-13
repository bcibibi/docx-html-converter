import { ImageRun, TextRun, TextWrappingSide, TextWrappingType, type IFloating, type IRunOptions, type XmlComponent } from "docx";
import type { Node } from "dom-parser";
import { NodeConverter, type GetChildrenFct } from "./node";
import type { ConverterContext } from "../context/convertercontext";
import debug from "debug";

const log = debug("docxhtml:converter:img");

export class IMGConverter extends NodeConverter {

  async convert(node: Node, run: IRunOptions, context: ConverterContext, children: GetChildrenFct): Promise<XmlComponent> {
    const src = node.getAttribute("src");
    const fileProvider = context.getFileProvider(src);
    const infos = await fileProvider.getFileInfos();
    if (!infos) {
      log(`No file infos retrieved for src: ${src}`);
      return new TextRun({ text: `[Image not found: ${src}]`, ...run });
    }
    const width = parseInt(node.getAttribute("width") || "100");
    const height = parseInt(node.getAttribute("height") || "100");
    const css = context.getCss(node);
    const floating = this.getFloatingValue(css);
    log(`Converting <img> node with src: ${src}, width: ${width}, height: ${height}, type: ${infos.type}, content size: ${infos.content.length} bytes`);
    return new ImageRun({
      type: infos.type,
      data: infos.content,
      transformation: {
        width,
        height
      },
      ...(floating ? { floating } : {})
    })
  }


  getFloatingValue(css: Record<string, string>): IFloating | undefined {
    const floatValue = css["float"];
    if (floatValue === "left" || floatValue === "right") {
      return {
        horizontalPosition: {
          relative: "column",
          align: floatValue === "left" ? "left" : "right"
        },
        verticalPosition: {
          relative: "paragraph",
          align: "top"
        },
        wrap: {
          type: TextWrappingType.SQUARE,
          side: TextWrappingSide.BOTH_SIDES,
        }
      };
    }
    return undefined;
  }

}