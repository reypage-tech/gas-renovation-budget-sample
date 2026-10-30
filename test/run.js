#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");
const vm = require("vm");

const codePath = path.join(__dirname, "..", "src", "Code.gs");
const raw = fs.readFileSync(codePath, "utf8");
const cut = raw.indexOf("function setupWorkbook");
const pureSource = cut === -1 ? raw : raw.slice(0, cut);

const sandbox = { module: { exports: {} }, console };
vm.createContext(sandbox);
vm.runInContext(pureSource, sandbox);

const {
  computeVariance_,
  monthlyTotals_,
  overrunLevel_,
  lowestBid_,
} = sandbox.module.exports;

let passed = 0;
let failed = 0;

function assert(cond, msg) {
  if (cond) {
    passed += 1;
    return;
  }
  failed += 1;
  console.error("FAIL:", msg);
}

assert(computeVariance_(100, 120) === 20, "variance basic");
assert(computeVariance_(null, 50) === 50, "variance null estimate");
assert(computeVariance_("", 0) === 0, "variance empty");

assert(overrunLevel_(100, 80) === 0, "overrun under 90");
assert(overrunLevel_(100, 91) === 1, "overrun warning band");
assert(overrunLevel_(100, 101) === 2, "overrun hard");
assert(overrunLevel_(0, 0) === 0, "overrun zero estimate");

assert(lowestBid_([12000, 11800, 12200]) === 11800, "lowest bid basic");
assert(lowestBid_([]) === null, "lowest bid empty");
assert(lowestBid_(["100", "", null, 90]) === 90, "lowest bid mixed");

const samplePayments = [
  { date: "2026-02-05", amount: 12000 },
  { date: "2026-02-28", amount: 3200 },
  { date: "2026-03-10", amount: 4500 },
];
const totals = monthlyTotals_(samplePayments);
assert(totals["2026-02"] === 15200, "monthly feb total");
assert(totals["2026-03"] === 4500, "monthly mar total");

assert(Object.keys(monthlyTotals_([])).length === 0, "monthly empty input");

const headerOnly = monthlyTotals_([{ date: "Date", amount: "Amount" }]);
assert(Object.keys(headerOnly).length === 0, "monthly header-only row");

const headerRowArray = monthlyTotals_([["Date", "Amount"]]);
assert(Object.keys(headerRowArray).length === 0, "monthly header-only array row");

console.log(`Tests: ${passed} passed, ${failed} failed`);
process.exit(failed === 0 ? 0 : 1);
