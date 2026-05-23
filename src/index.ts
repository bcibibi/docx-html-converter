
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

const log = debug("docxhtml:converter");

export namespace DocxHtmlConverter {

  export interface DocxHtmlConverterOptions {
    numbering?: string;
  }

  const nodeConverters: Record<string, NodeConverter> = {
    "#text": new TextConverter(),
    "p": new PConverter(),
    "span": new SPANConverter(),
    "ul": new ULConverter(),
    "ol": new OLConverter(),
    "li": new LIConverter(),
  };


  function parseHtml(html: string): Node | undefined {
    const doc = parseFromString(`<body>${html}</body>`);
    log(`Parsed HTML document with root node: ${doc.rawHTML}`);
    const body = doc.getElementsByTagName("body");
    return body.length > 0 ? body[0] : undefined;
  }

  export function convert(html: string, options?: DocxHtmlConverterOptions): FileChild[] {
    const result: FileChild[] = [];
    const context = new ConverterContext(options?.numbering || "");
    const node = parseHtml(html);
    log(`Parsed HTML with root node: ${node?.nodeName}`);
    for (const child of node?.childNodes || []) {
      const converted = convertNode(child, context);
      if (converted) {
        result.push(...converted.filter(c => c instanceof FileChild) as FileChild[]);
      } 
    }
    return result;
  }

  function convertNode(node: Node, context: ConverterContext, run: IRunOptions = {}): XmlComponent[] | undefined {
    log(`Converting node: ${node.nodeName}`);
    run = {...run, ...CSSParser.parse(node)};
    log(`Computed run options for node ${node.nodeName}: ${JSON.stringify(run)}`);
    const children = node.childNodes.map(child => convertNode(child, context, run)).filter(c => c !== undefined).flat() as XmlComponent[];
    log(`Converted node: ${node.nodeName} with ${children.length} children`);
    const converted = nodeConverters[node.nodeName]?.convert(node, run, children, context);
    return converted ? (Array.isArray(converted) ? converted : [converted]) : undefined;
  }

}