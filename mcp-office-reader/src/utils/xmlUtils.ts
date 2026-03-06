import { XMLParser } from "fast-xml-parser";

export const defaultXmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  textNodeName: "#text",
  isArray: (name) => {
    // Elements that should always be parsed as arrays
    const arrayElements = [
      "w:p",
      "w:r",
      "w:t",
      "w:tbl",
      "w:tr",
      "w:tc",
      "a:t",
      "a:r",
      "a:p",
      "p:sp",
      "Relationship",
    ];
    return arrayElements.includes(name);
  },
});

export function parseXml(xmlString: string): Record<string, unknown> {
  return defaultXmlParser.parse(xmlString) as Record<string, unknown>;
}

export function extractTextFromXmlNode(node: unknown): string {
  if (node === null || node === undefined) return "";
  if (typeof node === "string") return node;
  if (typeof node === "number" || typeof node === "boolean")
    return String(node);

  if (Array.isArray(node)) {
    return node.map(extractTextFromXmlNode).join("");
  }

  if (typeof node === "object") {
    const obj = node as Record<string, unknown>;
    // Handle text node
    if ("#text" in obj) return String(obj["#text"]);

    // Collect text from child nodes
    let text = "";
    for (const [key, value] of Object.entries(obj)) {
      if (!key.startsWith("@_")) {
        text += extractTextFromXmlNode(value);
      }
    }
    return text;
  }

  return "";
}

export function getNestedValue(
  obj: Record<string, unknown>,
  path: string
): unknown {
  const parts = path.split(".");
  let current: unknown = obj;
  for (const part of parts) {
    if (current === null || current === undefined) return undefined;
    if (typeof current !== "object") return undefined;
    current = (current as Record<string, unknown>)[part];
  }
  return current;
}

export function ensureArray<T>(value: T | T[] | undefined | null): T[] {
  if (value === null || value === undefined) return [];
  if (Array.isArray(value)) return value;
  return [value];
}
