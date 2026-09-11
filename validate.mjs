import { readFileSync } from "node:fs";
import { join } from "node:path";

const root = new URL(".", import.meta.url).pathname;
const read = (name) => readFileSync(join(root, name), "utf8");
const requiredFiles = [
  "Code.gs",
  "Data.gs",
  "SUMFONT.gs",
  "BRO.html",
  "Styles.html",
  "Assets.html",
  "Client.html",
  "appsscript.json"
];

for (const file of requiredFiles) read(file);

const bro = read("BRO.html");
const styles = read("Styles.html");
const client = read("Client.html");
const data = read("Data.gs");
const sumfont = read("SUMFONT.gs");

const requiredIds = [
  "focusButton",
  "focusDrawer",
  "presentationButton",
  "qualityButton",
  "copyReportButton",
  "copyFilterLinkButton",
  "registryBody"
];

for (const id of requiredIds) {
  if (!bro.includes(`id="${id}"`)) throw new Error(`Missing required element #${id}`);
}

const ids = [...bro.matchAll(/\bid="([^"]+)"/g)].map((match) => match[1]);
const duplicates = ids.filter((id, index) => ids.indexOf(id) !== index);
if (duplicates.length) throw new Error(`Duplicate IDs: ${[...new Set(duplicates)].join(", ")}`);

const staticIds = new Set(ids);
const dynamicIds = new Set(["annualProgressContracting"]);
const referencedIds = [...client.matchAll(/byId\("([^"]+)"\)/g)].map((match) => match[1]);
const missingIds = [...new Set(referencedIds)].filter((id) => !staticIds.has(id) && !dynamicIds.has(id));
if (missingIds.length) throw new Error(`Client references missing IDs: ${missingIds.join(", ")}`);

for (const include of ["Assets", "Styles", "Client"]) {
  if (!bro.includes(`include('${include}')`)) throw new Error(`BRO.html does not include ${include}.html`);
}

if (/data-theme-option="forest"|Тёмно-зелёная тема/.test(bro + styles + client)) {
  throw new Error("The rejected green theme is still present");
}

if (!data.includes("readContracting2025Total_")) {
  throw new Error("Structural 2025 contracting lookup is missing");
}

if (!sumfont.includes("function setupSUMFONT()") || !sumfont.includes("function SUMFONT_WATCH(e)")) {
  throw new Error("SUMFONT setup or watcher is missing");
}

const clientSource = client.replace(/^\s*<script>\s*/, "").replace(/\s*<\/script>\s*$/, "");
new Function(clientSource);
new Function(read("Code.gs"));
new Function(data);
new Function(sumfont);
JSON.parse(read("appsscript.json"));

console.log(`Validated ${requiredFiles.length} Apps Script files, ${ids.length} unique DOM IDs and ${referencedIds.length} client references.`);
