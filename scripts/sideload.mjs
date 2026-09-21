import { copyFile, link, mkdir, readFile, unlink } from "node:fs/promises";
import { existsSync } from "node:fs";
import os from "node:os";
import path from "node:path";

const candidates = [process.argv[2], "dist/manifest.xml", "manifest.xml"].filter(Boolean);
const source = candidates.find((candidate) => existsSync(candidate));

if (!source) {
  console.error("未找到 manifest。请先运行 npm run build，或指定 manifest 路径。");
  process.exit(1);
}

const manifest = await readFile(source, "utf8");
const idMatch = manifest.match(/<Id>\s*([^<\s]+)\s*<\/Id>/i);
if (!idMatch) {
  console.error("无法从清单中解析 <Id>，无法完成侧载。");
  process.exit(1);
}

const addinId = idMatch[1];
const basename = path.basename(source);
const destinationDir = path.join(
  os.homedir(),
  "Library/Containers/com.microsoft.Excel/Data/Documents/wef"
);

await mkdir(destinationDir, { recursive: true });

const plainPath = path.join(destinationDir, "manifest.xml");
const linkedPath = path.join(destinationDir, `${addinId}.${basename}`);

await copyFile(source, plainPath);

// Office 官方工具在 macOS 上使用 <GUID>.<文件名> 的硬链接登记加载项；
// 只放 manifest.xml 时功能区命令可能不注册，因此两种命名都写入。
try {
  if (existsSync(linkedPath)) {
    await unlink(linkedPath);
  }
  await link(plainPath, linkedPath);
} catch {
  await copyFile(source, linkedPath);
}

console.log(`已侧载 ${source}`);
console.log(`  - ${plainPath}`);
console.log(`  - ${linkedPath}（Office 工具同款命名）`);
console.log("若 Excel 正在运行，请完全退出并重新打开，然后在“插入/开始 → 加载项 → 我的加载项”中确认。");
