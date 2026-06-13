import { test } from "@jest/globals";
import { readFileSync, writeFileSync } from "fs";
import { DocxHtmlConverter } from "../src/index";
import { AlignmentType, convertInchesToTwip, Document, LevelFormat, Packer } from "docx";
import open from "open";

test("docxhtml", async () => {
  const numbering = "my-numbering";
  const html = readFileSync("./test/assets/index.html", "utf-8");
  const components = await DocxHtmlConverter.convert(html, { numbering });
  const document = new Document({
    sections: [
      {
        children: components,
      },
    ],
    numbering: {
      config: [
        {
          reference: numbering,
          levels: [
            {
              level: 0,
              format: "decimal",
              text: "%1.",
              alignment: "left",
              style: {
                paragraph: {
                  indent: { left: convertInchesToTwip(0.5), hanging: convertInchesToTwip(0.25) },
                },
              },
            },
            {
              level: 1,
              format: "decimal",
              text: "%1.%2.",
              alignment: "left",
              style: {
                paragraph: {
                  indent: { left: convertInchesToTwip(1), hanging: convertInchesToTwip(0.25) },
                },
              },
            },
            {
              level: 2,
              format: "decimal",
              text: "%1.%2.%3.",
              alignment: "left",
              style: {
                paragraph: {
                  indent: { left: convertInchesToTwip(1.5), hanging: convertInchesToTwip(0.25) },
                },
              },
            },
            {
              level: 3,
              format: "decimal",
              text: "%1.%2.%3.%4.",
              alignment: "left",
              style: {
                paragraph: {
                  indent: { left: convertInchesToTwip(2.5), hanging: convertInchesToTwip(0.25) },
                },
              },

            },
          ],
        },
      ]
    }
  });
  const buffer = await Packer.toBuffer(document);
  const filepath = "./test/assets/index.docx";
  writeFileSync(filepath, buffer);
  open(filepath);
}, 10000);
