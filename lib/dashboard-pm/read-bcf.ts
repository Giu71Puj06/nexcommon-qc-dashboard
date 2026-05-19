import JSZip from "jszip";

export type BCFMarkupRaw = {
  fileName: string;
  rawXml: string;
};

export async function readBCF(file: File): Promise<BCFMarkupRaw[]> {
  const zip = await JSZip.loadAsync(file);
  const markups: BCFMarkupRaw[] = [];

  for (const fileName of Object.keys(zip.files)) {
    const zipEntry = zip.files[fileName];

    if (!zipEntry.dir && fileName.toLowerCase().endsWith("markup.bcf")) {
      const rawXml = await zipEntry.async("text");

      markups.push({
        fileName,
        rawXml,
      });
    }
  }

  return markups;
}
