import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";

const DEV_ORIGIN = "https://localhost:3000";

const input = process.argv[2];
if (!input) {
  console.error("用法: npm run manifest:prod -- https://<用户名>.github.io/<仓库名>");
  process.exit(1);
}

let parsed;
try {
  parsed = new URL(input);
} catch {
  console.error(`无效的 URL：${input}`);
  process.exit(1);
}

if (parsed.protocol !== "https:") {
  console.error("Office 加载项必须使用 HTTPS 地址。");
  process.exit(1);
}

const baseUrl = `${parsed.origin}${parsed.pathname.replace(/\/+$/, "")}`;
const manifest = await readFile("manifest.xml", "utf8");

let output = manifest.split(DEV_ORIGIN).join(baseUrl);
output = output.replace(
  /<AppDomain>[^<]*<\/AppDomain>/,
  `<AppDomain>${parsed.origin}</AppDomain>`
);

// 给页面 URL 加上清单版本号，避免 Office 内置浏览器命中旧缓存导致更新不生效。
const version = output.match(/<Version>\s*([^<]+?)\s*<\/Version>/i)?.[1];
if (version) {
  output = output.replace(/taskpane\.html(?!\?)/g, `taskpane.html?v=${version}`);
}

await mkdir("dist", { recursive: true });
await writeFile(path.join("dist", "manifest.xml"), output, "utf8");

console.log(`已生成 dist/manifest.xml`);
console.log(`资源地址: ${baseUrl}`);
console.log(`AppDomain: ${parsed.origin}`);
