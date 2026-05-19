import JSZip from "jszip";

export async function readBCF(file: File) {
  const zip = await JSZip.loadAsync(file);

  const issues: any[] = [];

  for (const fileName of Object.keys(zip.files)) {
    if (fileName.endsWith("markup.bcf")) {
      const content = await zip.files[fileName].async("text");

      issues.push({
        file: fileName,
        raw: content,
      });
    }
  }

  return issues;
}
