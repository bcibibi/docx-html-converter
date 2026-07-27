
import { FileChild, XmlComponent, type IRunOptions } from "docx";
import { Node, parseFromString } from "dom-parser";
import { CSSParser } from "./css/cssparser.js";
import type { NodeConverter } from "./converter/node.js";
import debug from "debug";
import { PConverter } from "./converter/p.js";
import { TextConverter } from "./converter/text.js";
import { SPANConverter } from "./converter/span.js";
import { ConverterContext } from "./context/convertercontext.js";
import { ULConverter } from "./converter/ul.js";
import { LIConverter } from "./converter/li.js";
import { OLConverter } from "./converter/ol.js";
import { BConverter } from "./converter/b.js";
import { IConverter } from "./converter/i.js";
import { UConverter } from "./converter/u.js";
import { STRIKEConverter } from "./converter/strike.js";
import { SUBConverter } from "./converter/sub.js";
import { SUPConverter } from "./converter/sup.js";
import type { DocxHtmlFileProviderConstructor } from "./file/DocxHtmlFileProvider.js";
import { IMGConverter } from "./converter/img.js";
import type { DocxHtmlFileReader } from "./types/file.js";
import { defaultFileReader } from "./file/reader.js";

const log = debug("docxhtml:converter");

export { DocxHtmlFileProvider } from "./file/DocxHtmlFileProvider.js";
export { DefaultDocxHtmlFileProvider } from "./file/DefaultDocxHtmlFileProvider.js";
export type { DocxHtmlFileProviderConstructor } from "./file/DocxHtmlFileProvider.js";
export type { DocxHtmlFileReader, DocxHtmlFileType } from "./types/file.js";

export namespace DocxHtmlConverter {

  export interface DocxHtmlConverterOptions {
    numbering?: string;
    options?: IRunOptions;
    fileProvider?: DocxHtmlFileProviderConstructor;
    fileReader?: DocxHtmlFileReader;
  }

  const nodeConverters: Record<string, NodeConverter> = {
    "#text": new TextConverter(),
    "p": new PConverter(),
    "span": new SPANConverter(),
    "ul": new ULConverter(),
    "ol": new OLConverter(),
    "li": new LIConverter(),
    "b": new BConverter(),
    "strong": new BConverter(),
    "i": new IConverter(),
    "em": new IConverter(),
    "u": new UConverter(),
    "strike": new STRIKEConverter(),
    "sub": new SUBConverter(),
    "sup": new SUPConverter(),
    "img": new IMGConverter()
  };


  function parseHtml(html: string): Node | undefined {
    const doc = parseFromString(`<body>${html}</body>`);
    log(`Parsed HTML document with root node: ${doc.rawHTML}`);
    const body = doc.getElementsByTagName("body");
    return body.length > 0 ? body[0] : undefined;
  }

  export async function convert(html: string, options?: DocxHtmlConverterOptions): Promise<FileChild[]> {
    const result: FileChild[] = [];
    const context = new ConverterContext(
      options?.numbering || "",
      options?.fileProvider,
      options?.fileReader || defaultFileReader
    );
    const node = parseHtml(html);
    log(`Parsed HTML with root node: ${node?.nodeName}`);
    for (const child of node?.childNodes || []) {
      const converted = await convertNode(child, context, options?.options || {});
      if (converted) {
        result.push(...converted.filter(c => c instanceof FileChild) as FileChild[]);
      }
    }
    return result;
  }

  async function convertNode(node: Node, context: ConverterContext, run: IRunOptions = {}): Promise<XmlComponent[] | undefined> {
    log(`Converting node: ${node.nodeName}`);
    run = { ...run, ...CSSParser.parse(node, context) };
    log(`Computed run options for node ${node.nodeName}: ${JSON.stringify(run)}`);
    const converted = await nodeConverters[node.nodeName]?.convert(node, run, context, convertChildren);
    return converted ? (Array.isArray(converted) ? converted : [converted]) : undefined;
  }

  async function convertChildren(node: Node, run: IRunOptions, context: ConverterContext): Promise<XmlComponent[]> {
    const children = await Promise.all(node.childNodes.map(child => convertNode(child, context, run)));
    return children.filter(c => c !== undefined).flat() as XmlComponent[];
  }

}