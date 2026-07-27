import { ImageRun, TextRun, TextWrappingSide, TextWrappingType, type IFloating, type IRunOptions, type XmlComponent } from "docx";
import type { Node } from "dom-parser";
import { NodeConverter, type GetChildrenFct } from "./node.js";
import type { ConverterContext } from "../context/convertercontext.js";
import debug from "debug";
import { FileUtils } from "../file/utils.js";
import type { DocxHtmlFileType } from "../types/file.js";

const log = debug("docxhtml:converter:img");

export class IMGConverter extends NodeConverter {

  async convert(node: Node, run: IRunOptions, context: ConverterContext, children: GetChildrenFct): Promise<XmlComponent> {
    const src = node.getAttribute("src");
    const infos = await this.getFile(context, src);
    if (!infos) {
      log(`No file content found for src: ${src}`);
      return new TextRun({
        text: `[Image not found: ${src}]`,
        ...run
      });
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

  private async getFile(context: ConverterContext, src: string): Promise<{ content: Buffer, type: DocxHtmlFileType } | undefined> {
    try {
      const fileProvider = context.getFileProvider(src);
      const content = fileProvider ? await fileProvider.getFileContent() : await context.fileReader(src);
      return {
        content,
        type: FileUtils.getFileType(content)
      }
    } catch (error) {
      log(`Error retrieving file for src: ${src} - ${error}`);
      return undefined;
    }
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