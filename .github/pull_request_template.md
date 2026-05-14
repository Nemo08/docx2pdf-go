<!--
Thanks for the PR! A few things make review fast:

  - One PR per logical change. Bundle "feature + regression test" but
    not "feature + unrelated cleanup".
  - The verify harness (internal/verify/verify_test.go) is the canonical
    place for regression cases — add one if you're touching the parser
    or renderer.
  - Run `gofmt -w .`, `go vet ./...`, `go test ./... -race` locally
    before pushing.
-->

## Summary

<!-- One or two sentences on what this changes and why. -->

## What this affects

- [ ] Parser (internal/docx)
- [ ] Renderer (internal/render)
- [ ] Public API (docx2pdf.go)
- [ ] CLI (cmd/docx2pdf)
- [ ] Docker image / fonts
- [ ] Tests / CI

## Verification

<!--
What did you do to confirm this works? Test names, sample docx, before/
after PNG screenshots welcome.
-->

## Notes for the reviewer

<!-- Anything surprising, or a design decision that deserves a heads-up. -->
