import { XMLParser } from "fast-xml-parser";

export type BCFIssue = {
  guid: string;
  title: string;
  creationDate: string;
  lastActivityDate: string;
  status: string;
  author: string;
  commentsCount: number;
  isClosed: boolean;
};

function toArray<T>(value: T | T[] | undefined): T[] {
  if (!value) return [];
  return Array.isArray(value) ? value : [value];
}

function normalizeStatus(status: string): string {
  return String(status || "").trim().toLowerCase();
}

function isClosedStatus(status: string): boolean {
  const s = normalizeStatus(status);

  return [
    "closed",
    "chiuso",
    "chiusa",
    "resolved",
    "risolto",
    "risolta",
    "done",
    "completed",
    "concluso",
    "conclusa",
  ].includes(s);
}

function maxDate(dates: string[]): string {
  const validDates = dates
    .filter(Boolean)
    .map((d) => new Date(d))
    .filter((d) => !Number.isNaN(d.getTime()));

  if (validDates.length === 0) return "";

  return new Date(
    Math.max(...validDates.map((d) => d.getTime()))
  ).toISOString();
}

export function extractBCFHistory(rawXml: string): BCFIssue | null {
  try {
    const parser = new XMLParser({
      ignoreAttributes: false,
    });

    const parsed = parser.parse(rawXml);
    const topic = parsed?.Markup?.Topic;

    if (!topic) return null;

    const comments = toArray(parsed?.Markup?.Comment);

    const commentDates = comments
      .map((comment: any) => comment.Date)
      .filter(Boolean);

    const creationDate = topic.CreationDate || "";
    const modifiedDate = topic.ModifiedDate || "";

    const lastActivityDate =
      maxDate([creationDate, modifiedDate, ...commentDates]) ||
      modifiedDate ||
      creationDate;

    const status =
      topic.TopicStatus ||
      topic.Status ||
      topic["@_TopicStatus"] ||
      "";

    return {
      guid: topic["@_Guid"] || topic.Guid || "",
      title: topic.Title || "",
      creationDate,
      lastActivityDate,
      status,
      author: topic.CreationAuthor || "",
      commentsCount: comments.length,
      isClosed: isClosedStatus(status),
    };
  } catch (error) {
    console.error("Errore parsing BCF:", error);
    return null;
  }
}
