import { readFile } from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { spawn } from "node:child_process";

const require = createRequire(import.meta.url);
const manifestLib = require("office-addin-manifest");
const devSettings = require("office-addin-dev-settings");

const manifestPath = path.resolve(process.argv[2] ?? "dist/manifest.xml");
const shouldOpen = process.argv.includes("--open");

const xml = await readFile(manifestPath, "utf8");
const id = xml.match(/<Id>\s*([^<]+?)\s*<\/Id>/i)?.[1];
const version = xml.match(/<Version>\s*([^<]+?)\s*<\/Version>/i)?.[1];
const officeAppType = xml.match(/xsi:type="([^"]+)"/i)?.[1] ?? "TaskPaneApp";

if (!id || !version) {
  console.error("无法从清单解析 <Id> 或 <Version>。");
  process.exit(1);
}

const manifest = {
  id,
  version,
  officeAppType,
  manifestType: manifestLib.ManifestType.XML
};

// Excel 只在“打开带有该加载项注册信息的文档”时才会注册功能区命令，
// 因此生成一个临时工作簿并（可选）打开它。
const documentPath = await devSettings.generateSideloadFile("excel", manifest);
console.log(`已生成注册用工作簿：${documentPath}`);

if (!shouldOpen) {
  console.log("如需注册加载项，请手动打开该文件，或使用 npm run register。");
  process.exit(0);
}

await new Promise((resolve, reject) => {
  const child = spawn("open", [documentPath], { stdio: "ignore" });
  child.on("exit", (code) => {
    if (code === 0) {
      resolve();
    } else {
      reject(new Error(`open 退出码 ${code}`));
    }
  });
});

console.log("已在 Excel 中打开注册用工作簿：看到“标签栏”按钮后即可关闭该临时文档。");
