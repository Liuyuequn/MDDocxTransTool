import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import mammoth from "mammoth";

export function createImageConverter(outputDir, pictureDirName) {
  const pictureDir = path.join(outputDir, pictureDirName);
  return mammoth.images.imgElement(async (image) => {
    const buffer = await image.read();
    const ext = extensionFromContentType(image.contentType);
    const hash = crypto.createHash("md5").update(buffer).digest("hex").slice(0, 10);
    const name = `img-${hash}.${ext}`;
    fs.mkdirSync(pictureDir, { recursive: true });
    fs.writeFileSync(path.join(pictureDir, name), buffer);
    return { src: `${pictureDirName}/${name}` };
  });
}

function extensionFromContentType(contentType) {
  const known = {
    "image/png": "png",
    "image/jpeg": "jpg",
    "image/gif": "gif",
    "image/bmp": "bmp",
  };
  if (known[contentType]) return known[contentType];
  const subtype = String(contentType).split("/")[1] || "img";
  return subtype.replace(/[^a-z0-9]/gi, "") || "img";
}
