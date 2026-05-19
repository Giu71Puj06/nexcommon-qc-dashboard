import { XMLParser } from "fast-xml-parser";

export type BCFIssue = {
  guid: string;
  title: string;
  creationDate: string;
  modifiedDate: string;
  status: string;
  author: string;
  commentsCount: number;
};

export function extractBCFHistory(rawXml: string): BCFIssue | null {
  try {
    const parser = new XMLParser({
      ignoreAttributes: false,
    });

    const parsed = parser.parse(rawXml);

    const topic = parsed?.Markup?.Topic;

    if (!topic) {
      return null;
    }

    const comments = parsed?.Markup?.Comment;

    let commentsCount = 0;

    if (Array.isArray(comments)) {
      commentsCount = comments.length;
    } else if (comments) {
      commentsCount = 1;
    }

    return {
      guid: topic["@_Guid"] || "",
      title: topic.Title || "",
      creationDate: topic.CreationDate || "",
      modifiedDate: topic.ModifiedDate || "",
      status: topic.TopicStatus || "",
      author: topic.CreationAuthor || "",
      commentsCount,
    };
  } catch (error) {
    console.error("Errore parsing BCF:", error);
    return null;
  }
}
