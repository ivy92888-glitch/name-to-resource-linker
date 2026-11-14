# Push guide

How this repository was published to GitHub. Kept here for reproducibility; no
secret is stored in this file.

## Prerequisites

- A GitHub account (the author used `ivy92888-glitch`).
- A Personal Access Token (classic) with `repo` scope.
- `curl` (or any HTTP client) able to reach `api.github.com`.

> Treat the token as a secret. Pass it through an environment variable, never
> hard-code it in a file, and rotate it after use.

## Steps

1. Create the repository (public, default branch `main`):

   ```bash
   curl -s -X POST \
     -H "Authorization: Bearer $GH_TOKEN" \
     -H "Content-Type: application/json" \
     -d '{"name":"name-to-resource-linker","description":"Generic free-text item name to ranked reference resource link matcher (VBA).","public":true,"auto_init":false}' \
     https://api.github.com/user/repos
   ```

2. Upload each file via the Contents API. For a new file use `PUT` with a base64
   body and a commit message; for an update you must also supply the blob `sha`.

   ```bash
   # example for README.md
   CONTENT=$(base64 -w0 README.md)
   curl -s -X PUT \
     -H "Authorization: Bearer $GH_TOKEN" \
     -H "Content-Type: application/json" \
     -d "{\"message\":\"Add README\",\"content\":\"$CONTENT\"}" \
     "https://api.github.com/repos/ivy92888-glitch/name-to-resource-linker/contents/README.md"
   ```

   Repeat for `LICENSE`, `src/NameToResource.bas`, `docs/Configuration.md`,
   `examples/sample-data.xlsx`, `.gitignore`, and `.gitattributes`.

3. Verify on `https://github.com/ivy92888-glitch/name-to-resource-linker`.

## Notes

- The macro workbook (`examples/sample-data.xlsx`) is committed as a binary asset;
  it is fictional sample data only.
- The repository intentionally contains no real data, internal names, or
  third-party logic.
- A desensitization scan was run over all repository text and the sample
  workbook; it reported zero hits for any proprietary or business-specific term.
