# Security Policy

## Supported versions

| Version | Supported |
| ------- | --------- |
| 2.x     | Yes       |
| < 2.0   | No — depends on the vulnerable `jszip@2` and abandoned `xmldom`; please upgrade |

## Reporting a vulnerability

Please report security issues privately via GitHub Security Advisories
("Report a vulnerability" on the repository's Security tab) rather than opening a
public issue. We aim to acknowledge within a few days and to disclose a fix
within 90 days of a confirmed report.

## Threat model

`docx-merger` parses `.docx` files, which are zip archives of XML. Inputs are
treated as **untrusted**: parsing goes through the maintained
[`@xmldom/xmldom`](https://github.com/xmldom/xmldom) and
[`fflate`](https://github.com/101arrowz/fflate), and the production dependency
tree is kept advisory-free (`npm audit` runs in CI on every change).

Note that merging executes no document content; it only rewrites OOXML parts.
Still, treat output opened in Microsoft Word as you would any document built
from untrusted input.
