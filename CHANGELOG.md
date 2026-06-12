# Changelog

## 2.0.0

A security and correctness release. The documented public API is unchanged:
`new DocxMerger(options, files)` is still synchronous, and
`save(type, callback)` still calls back synchronously. Your existing code keeps
working without changes.

### Security

- Removed `jszip@2` (prototype pollution GHSA-jg8v-48h5-wgxg, path traversal
  GHSA-36fh-84j7-cv5h) — replaced internally with `fflate`.
- Removed the abandoned `xmldom` (8 advisories incl. critical
  GHSA-crh6-fp67-6883) — replaced with the maintained `@xmldom/xmldom@0.8`.
- `npm audit` is now clean (was 1 critical + 1 moderate) and enforced in CI.

### Fixed

- **Output corruption** when document text contained `$&`, `` $` ``, `$'`:
  `String.replace` interpreted them as replacement patterns and duplicated or
  mangled content. All section rewrites now splice by index. (#48, #18)
- **Crash / silent skip** when a style or relationship ID contained regex
  metacharacters — IDs are now escaped before being used in a `RegExp`.
- **Dropped image content type** when two `<Default>` entries shared a
  ContentType (e.g. `jpg` and `jpeg` both `image/jpeg`). Dedupe is now keyed on
  Extension / PartName. (#53)
- **Duplicated / colliding media** on names like `image10.png`; media is now
  renamed to a collision-free `media_N.ext` and the base file's originals are no
  longer left as orphans. (#55)
- **Silently dropped hyperlinks** when two files reused the same relationship ID
  for different targets — non-base hyperlink IDs are renamed per file.
- **Broken list numbering** ("Word found unreadable content" repair prompt):
  numbering IDs are globally renumbered and the body's `w:numId` references are
  rewritten to match (v1 left them dangling). (#56, #43)
- Body is no longer truncated when `<w:sectPr>` is absent.

### Added

- `save(type)` with no callback returns a `Promise`.
- Descriptive, indexed errors for invalid or incomplete `.docx` inputs (was an
  opaque `TypeError`).
- Hand-written TypeScript definitions (`types/index.d.ts`).

### Performance

- `document.xml` is rewritten once per file instead of once per style and once
  per media file. A 2×(300 styles, 2000 paragraphs) merge dropped from
  ~600–900 ms to ~120 ms.

### Build / tooling

- Browser bundle is built with esbuild (~38 kB gzipped) and exposes the same
  global `DocxMerger`.
- Removed the dead Babel 6+7 / webpack 4 / Travis / semantic-release toolchain
  (~1600 dev packages). Tests run on Node's built-in test runner. CI runs on
  Node 18/20/22 across Linux and Windows.

### Observable changes that justify the major bump

These do not affect the documented API, but output bytes differ from 1.x:

- Merged media parts are named `media_N.ext` (were `image_N.ext`, often with
  orphaned duplicates).
- Numbering IDs are renumbered; rendered documents are identical or repair-free.
- The constructor throws a descriptive `Error` (not an opaque `TypeError`) on
  invalid input.
- The `dist/` bundle is produced by esbuild; the script-tag global is unchanged.
