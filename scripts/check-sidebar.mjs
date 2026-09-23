import { readFile } from "node:fs/promises";
import vm from "node:vm";

const html = await readFile(new URL("../Sidebar.html", import.meta.url), "utf8");
const scripts = [...html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/gi)];

if (scripts.length === 0) {
  throw new Error("Sidebar.html contains no inline script to validate");
}

for (const [index, match] of scripts.entries()) {
  new vm.Script(match[1], { filename: `Sidebar.html#script-${index + 1}` });
}

console.log(`Validated ${scripts.length} inline Sidebar.html script block(s).`);
