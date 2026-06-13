
import { FileChild, XmlComponent, type IRunOptions } from "docx";
import { Node, parseFromString } from "dom-parser";
import { CSSParser } from "./css/cssparser";
import type { NodeConverter } from "./converter/node";
import debug from "debug";
import { PConverter } from "./converter/p";
import { TextConverter } from "./converter/text";
import { SPANConverter } from "./converter/span";
import { ConverterContext } from "./context/convertercontext";
import { ULConverter } from "./converter/ul";
import { LIConverter } from "./converter/li";
import { OLConverter } from "./converter/ol";
import { BConverter } from "./converter/b";
import { IConverter } from "./converter/i";
import { UConverter } from "./converter/u";
import { STRIKEConverter } from "./converter/strike";
import { SUBConverter } from "./converter/sub";
import { SUPConverter } from "./converter/sup";
import type { DocxHtmlFileProviderConstructor } from "./file/DocxHtmlFileProvider";
import { DefaultDocxHtmlFileProvider } from "./file/DefaultDocxHtmlFileProvider";
import { IMGConverter } from "./converter/img";

const log = debug("docxhtml:converter");

export { DocxHtmlFileProvider } from "./file/DocxHtmlFileProvider";
export { DefaultDocxHtmlFileProvider } from "./file/DefaultDocxHtmlFileProvider";

export namespace DocxHtmlConverter {

  export interface DocxHtmlConverterOptions {
    numbering?: string;
    fileProvider?: DocxHtmlFileProviderConstructor;
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
    const context = new ConverterContext(options?.numbering || "", options?.fileProvider || DefaultDocxHtmlFileProvider);
    const node = parseHtml(html);
    log(`Parsed HTML with root node: ${node?.nodeName}`);
    for (const child of node?.childNodes || []) {
      const converted = await convertNode(child, context);
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