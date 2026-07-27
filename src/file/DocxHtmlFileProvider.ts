import debug from "debug";

const log = debug("docxhtml:fileprovider");

/**
 * @deprecated Use 'fileReader' instead of 'fileProvider' in DocxHtmlConverterOptions. This class will be removed in future versions.
 */
export type DocxHtmlFileProviderConstructor<T extends DocxHtmlFileProvider = DocxHtmlFileProvider> = new (src: string) => T;

/**
 * @deprecated Use 'fileReader' instead of 'fileProvider' in DocxHtmlConverterOptions. This class will be removed in future versions.
 */
export abstract class DocxHtmlFileProvider {

    constructor(protected readonly src: string) {}

    abstract getFileContent(): Promise<Buffer>;

}