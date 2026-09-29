import { readFile } from "node:fs/promises";

const path = new URL("../appsscript.json", import.meta.url);
const manifest = JSON.parse(await readFile(path, "utf8"));

if (manifest.runtimeVersion !== "V8") {
  throw new Error("appsscript.json must use the V8 runtime");
}

if (typeof manifest.timeZone !== "string" || manifest.timeZone.length === 0) {
  throw new Error("appsscript.json must declare a timeZone");
}

console.log("Validated appsscript.json.");
