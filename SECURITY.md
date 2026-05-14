# Security Policy

## Supported versions

docx2pdf-go is a single-binary library. Security fixes are applied to
the latest minor release line; older versions are not maintained.

## Reporting a vulnerability

**Please do NOT open a public GitHub issue** for vulnerabilities.

Instead, use GitHub's private vulnerability reporting:

1. Open [Security tab](https://github.com/bobyeoh/docx2pdf-go/security)
   of the repository
2. Click *Report a vulnerability*

We aim to acknowledge reports within 72 hours and ship a fix within 30
days for confirmed issues. You'll be credited in the release notes
unless you prefer to stay anonymous.

## Threat model

docx2pdf-go parses untrusted `.docx` input — a zipped XML bundle from
the network or a user upload. We treat this as adversarial input. The
test suite includes crash-resistance cases covering:

- Empty / truncated zips
- Malformed XML (mismatched tags, encoding violations)
- Circular `basedOn` style chains
- Deeply nested elements (500+ levels)
- Corrupt image bytes

Known limitations:

- No specific defense against zip-bombs (decompression-ratio attacks).
  Wrap calls with `convert.ConvertContext` and a strict deadline if
  you accept untrusted input.
- No sandboxing of font loading — a maliciously crafted TTF could
  exploit issues in the upstream `signintech/gopdf` parser. We do not
  control that surface; vulnerabilities there should be reported
  upstream.

If you find something we haven't covered, we'd like to know.
