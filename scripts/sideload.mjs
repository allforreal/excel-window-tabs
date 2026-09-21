import { copyFile, mkdir } from "node:fs/promises";
import { existsSync } from "node:fs";
import os from "node:os";
import path from "node:path";

const candidates = [process.argv[2], "dist/manifest.xml", "manifest.xml"].filter(Boolean);
const source = candidates.find((candidate) => existsSync(candidate));

if (!source) {
  console.error("未找到 manifest。请先运行 npm run build，或指定 manifest 路径。");
  process.exit(1);
}

const destinationDir = path.join(
  os.homedir(),
  "Library/Containers/com.microsoft.Excel/Data/Documents/wef"
);
const destination = path.join(destinationDir, path.basename(source));

await mkdir(destinationDir, { recursive: true });
await copyFile(source, destination);

console.log(`已侧载 ${source} -> ${destination}`);
console.log("若 Excel 正在运行，请完全退出并重新打开，然后在“插入 → 我的加载项”中确认。");
