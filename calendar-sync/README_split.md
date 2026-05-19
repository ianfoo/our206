# Our206 Apps Script split

This package splits the refactored single-file Apps Script into multiple `.js` files for use with `clasp`.

Apps Script still uses one shared global namespace, so these files are organizational only; top-level function and constant names must remain unique.

Suggested use:

```bash
unzip our206_split.zip -d our206_split
cp our206_split/*.js /path/to/your/clasp/project/
clasp push
```

If your existing project already has `appsscript.json`, keep your existing manifest.
