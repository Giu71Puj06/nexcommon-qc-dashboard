import JSZip from "jszip";

type BCFTopic = {
  title: string;
  description: string;
  author?: string;
  status?: string;
};

function escapeXml(value: string) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function createGuid() {
  return crypto.randomUUID();
}

export async function generateBCFZip(topics: BCFTopic[]) {
  const zip = new JSZip();

  zip.file(
    "bcf.version",
    `<?xml version="1.0" encoding="UTF-8"?>
<Version VersionId="2.1" />`
  );

  topics.forEach((topic) => {
    const guid = createGuid();
    const now = new Date().toISOString();

    const markup = `<?xml version="1.0" encoding="UTF-8"?>
<Markup>
  <Topic Guid="${guid}" TopicType="Issue" TopicStatus="${escapeXml(
      topic.status || "Open"
    )}">
    <Title>${escapeXml(topic.title)}</Title>
    <CreationDate>${now}</CreationDate>
    <CreationAuthor>${escapeXml(topic.author || "Dashboard PM")}</CreationAuthor>
    <Description>${escapeXml(topic.description)}</Description>
  </Topic>
  <Comment Guid="${createGuid()}">
    <Date>${now}</Date>
    <Author>${escapeXml(topic.author || "Dashboard PM")}</Author>
    <Comment>${escapeXml(topic.description)}</Comment>
    <Topic Guid="${guid}" />
  </Comment>
</Markup>`;

    zip.folder(guid)?.file("markup.bcf", markup);
  });

  return zip.generateAsync({
    type: "blob",
    mimeType: "application/zip",
  });
}

export function downloadBlob(blob: Blob, fileName: string) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = fileName;
  link.click();

  URL.revokeObjectURL(url);
}
