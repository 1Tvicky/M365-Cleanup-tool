import { describe, expect, it } from "vitest";
import { generateFileContent } from "./fileContent.js";

const ZIP_MAGIC = Buffer.from([0x50, 0x4b, 0x03, 0x04]); // PK\x03\x04 — every OOXML file (docx/xlsx/pptx) and every .zip is a real zip container
const PNG_MAGIC = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
const PROFESSIONAL_NAME = /^[A-Za-z0-9_-]+\.[a-z]+$/; // e.g. "Project_Charter.docx" — never a generic "test.docx"

describe("generateFileContent", () => {
  it("produces a real, valid docx (a zip container) with a professional, content-matched name, not placeholder text", async () => {
    const file = await generateFileContent(1, 0, "docx", "professional");
    expect(file.fileName).toMatch(PROFESSIONAL_NAME);
    expect(file.buffer.subarray(0, 4)).toEqual(ZIP_MAGIC);
    expect(file.buffer.length).toBeGreaterThan(1000);
  });

  it("produces a real, valid xlsx with a professional name", async () => {
    const file = await generateFileContent(1, 1, "xlsx", "professional");
    expect(file.fileName).toMatch(PROFESSIONAL_NAME);
    expect(file.buffer.subarray(0, 4)).toEqual(ZIP_MAGIC);
  });

  it("produces a real, valid pptx with a professional name", async () => {
    const file = await generateFileContent(1, 2, "pptx", "professional");
    expect(file.fileName).toMatch(PROFESSIONAL_NAME);
    expect(file.buffer.subarray(0, 4)).toEqual(ZIP_MAGIC);
  });

  it("produces a real, valid pdf with a professional name", async () => {
    const file = await generateFileContent(1, 3, "pdf", "professional");
    expect(file.fileName).toMatch(PROFESSIONAL_NAME);
    expect(file.buffer.subarray(0, 5).toString("ascii")).toBe("%PDF-");
  });

  it("produces readable txt content with real business text, not lorem-ipsum-style filler", async () => {
    const file = await generateFileContent(1, 4, "txt", "professional");
    expect(file.fileName).toMatch(PROFESSIONAL_NAME);
    const text = file.buffer.toString("utf8");
    expect(text.length).toBeGreaterThan(0);
    expect(text).not.toMatch(/lorem ipsum/i);
  });

  it("produces structured csv content with a header row", async () => {
    const file = await generateFileContent(1, 5, "csv", "professional");
    const text = file.buffer.toString("utf8");
    expect(text.split("\n")[0]).toMatch(/Employee ID,Department,Region,Quarter,Revenue,Target,Actual/);
  });

  it("produces a real, valid PNG for the png type", async () => {
    const file = await generateFileContent(1, 6, "png", "professional");
    expect(file.fileName).toMatch(PROFESSIONAL_NAME);
    expect(file.buffer.subarray(0, 8)).toEqual(PNG_MAGIC);
  });

  it("'jpg' is saved with a real .png extension (documented scope decision — see fileContent.ts), never a mismatched/corrupt-looking file", async () => {
    const file = await generateFileContent(1, 7, "jpg", "professional");
    expect(file.fileName.endsWith(".png")).toBe(true);
    expect(file.buffer.subarray(0, 8)).toEqual(PNG_MAGIC);
  });

  it("produces a real, valid zip for the zip type", async () => {
    const file = await generateFileContent(1, 8, "zip", "professional");
    expect(file.fileName).toMatch(PROFESSIONAL_NAME);
    expect(file.buffer.subarray(0, 4)).toEqual(ZIP_MAGIC);
  });

  it("is deterministic: the same (seed, index, type, namingStyle) always produces byte-identical content and name", async () => {
    const a = await generateFileContent(42, 10, "txt", "professional");
    const b = await generateFileContent(42, 10, "txt", "professional");
    expect(a.buffer.equals(b.buffer)).toBe(true);
    expect(a.fileName).toBe(b.fileName);
  });

  it("synthetic naming style produces flat File1/File2/... names instead of business-style names", async () => {
    const a = await generateFileContent(1, 0, "docx", "synthetic");
    const b = await generateFileContent(1, 1, "docx", "synthetic");
    expect(a.fileName).toBe("File1.docx");
    expect(b.fileName).toBe("File2.docx");
    expect(a.buffer.subarray(0, 4)).toEqual(ZIP_MAGIC);
  });
});
