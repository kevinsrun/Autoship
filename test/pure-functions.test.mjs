import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";

async function loadCode() {
  const source = await readFile(new URL("../Code.js", import.meta.url), "utf8");
  const context = vm.createContext({ console, Date, Set });
  new vm.Script(source, { filename: "Code.js" }).runInContext(context);
  return context;
}

test("scholarshipHit_ recognizes a scholarship signal", async () => {
  const code = await loadCode();
  assert.equal(code.scholarshipHit_("Community Scholarship", ""), "contains:scholarship");
});

test("scholarshipHit_ rejects unrelated content", async () => {
  const code = await loadCode();
  assert.equal(code.scholarshipHit_("Weekly lunch menu", "Pizza on Friday"), "");
});

test("stripHtml_ removes executable and presentation markup", async () => {
  const code = await loadCode();
  assert.equal(code.stripHtml_("<style>x</style><p>Hello <strong>student</strong></p><script>bad()</script>"), "Hello student");
});

test("pickCourses_ requires every configured matcher", async () => {
  const code = await loadCode();
  const courses = [
    { id: 1, name: "Class of 2026" },
    { id: 2, name: "Class of 2025" },
    { id: 3, name: "Counselors 2026" }
  ];
  const result = code.pickCourses_(courses, [/class/i, /2026/i]);
  assert.deepEqual(result.map((course) => course.id), [1]);
});
