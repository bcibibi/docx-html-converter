import debug from "debug";

const log = debug("docxhtml:fileprovider");

export type DocxHtmlFileType = "jpg" | "png" | "gif" | "bmp";

export type DocxHtmlFileProviderConstructor<T extends DocxHtmlFileProvider = DocxHtmlFileProvider> = new (src: string) => T;

export abstract class DocxHtmlFileProvider {

    constructor(protected readonly src: string) {}

    abstract getFileContent(): Promise<Buffer>;

    getFileType(content: Buffer): DocxHtmlFileType {
        const header = content.subarray(0, 4).toString("hex");
        log(`Determining file type from content header: ${header}`);
        switch (header) {
            case "ffd8ffe0":
            case "ffd8ffe1":
            case "ffd8ffe2":
                return "jpg";
            case "89504e47":
                return "png";
            case "47494638":
                return "gif";
            case "424d":
                return "bmp";
            default:
                throw new Error(`Unsupported file type from content: ${header}`);
        }
    }

    async getFileInfos(): Promise<{ content: Buffer, type: DocxHtmlFileType } | undefined> {
        try {
            const content = await this.getFileContent();
            const type = this.getFileType(content);
            log(`Retrieved file content of type ${type} with size ${content.length} bytes`);
            return { content, type };
        } catch (error) {
            log(`Error retrieving file content for src ${this.src}: ${error}`);
            console.error(`Error retrieving file content for src ${this.src}:`, error);
        }
    }

}