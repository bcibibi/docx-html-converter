import debug from "debug";
import type { DocxHtmlFileType } from "../types/file.js";

const log = debug("docxhtml:fileutils");


export namespace FileUtils {

    export function getFileType(content: Buffer): DocxHtmlFileType {
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

}