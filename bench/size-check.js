'use strict';
// Regression guard on shipped size. Budgets are set with headroom over the
// current footprint; they exist to catch a dependency or bundle blow-up, not to
// enforce an arbitrary number. Gzipped is what a browser actually downloads, so
// that is the budget that matters for the bundle.
var fs = require('fs');
var zlib = require('zlib');
var execSync = require('child_process').execSync;

var BUNDLE = 'dist/docx-merger.min.js';
var GZIP_BUDGET_KB = 50;   // current ~38 kB
var TARBALL_BUDGET_KB = 70; // current ~49 kB packed

var failures = [];

if (!fs.existsSync(BUNDLE)) {
    console.error('FAIL: ' + BUNDLE + ' not found - run `npm run build` first');
    process.exit(1);
}

var raw = fs.readFileSync(BUNDLE);
var gzipKb = zlib.gzipSync(raw).length / 1024;
console.log('browser bundle: ' + (raw.length / 1024).toFixed(1) + ' kB raw, ' +
            gzipKb.toFixed(1) + ' kB gzipped (budget ' + GZIP_BUDGET_KB + ' kB gzipped)');
if (gzipKb > GZIP_BUDGET_KB) failures.push('bundle over gzip budget');

var meta = JSON.parse(execSync('npm pack --dry-run --json').toString())[0];
var tarballKb = meta.size / 1024;
console.log('npm tarball: ' + tarballKb.toFixed(1) + ' kB packed, ' + meta.entryCount + ' files (budget ' + TARBALL_BUDGET_KB + ' kB)');
if (tarballKb > TARBALL_BUDGET_KB) failures.push('tarball over budget');

if (failures.length) {
    console.error('FAIL: ' + failures.join('; '));
    process.exit(1);
}
console.log('size OK');
