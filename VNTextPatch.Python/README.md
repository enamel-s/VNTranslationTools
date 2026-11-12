# vntextpatch.py

A Python port of VNTextPatch's Ren'Py workflow that reads scripts straight from
`.rpa` archives, writes the same `template.xlsx`-style workbook, patches the
translation back, and can fill the workbook with a Gemini translation
(Burmese by default).

Requires Python 3.8+ and `pip install openpyxl`.

## Menu

Run `python vntextpatch.py` with no arguments for an interactive menu: extract,
build the translation context, test run on a few lines, translate, show
progress per sheet, insert, and settings (API keys, main model and fallbacks,
batch size, language, extraction mode, extra patterns). Paths and settings you enter are remembered in
`.vntextpatch.json` and the Gemini ones in `gemini_config.json` in the current folder; pasted API keys are kept in memory
only. Everything in the menu is also available as the commands below.

## 1. Extract

```
python vntextpatch.py extract <game folder | scripts.rpa | folder of .rpy> script.xlsx
```

* `.rpa` archives (RPA-2.0/3.0) found in the input are read directly; loose
  `.rpy` files win over archived ones. Archives that only hold compiled `.rpyc`
  need [unrpyc](https://github.com/CensoredUsername/unrpyc): add
  `--unrpyc path/to/unrpyc.py`.
* One sheet per script, columns as in VNTextPatch: original character,
  original line, translated character, `TL`, `TLC`, `Edit`, `Notes`.
  `Notes` is prefilled with the line type (`say`, `choice`, `name`, `string`).
* `names.xml` (same format as VNTextPatch) is written next to the workbook.
* `--mode dialogue` (default) extracts say statements, menu choices, character
  names, `_("...")` strings and screen texts. `--mode all` extracts every string
  literal exactly like `RenpyScript.cs` (this includes file names).
* `--extra REGEX` extracts more strings: the regex is matched against the text
  before the opening quote on that line, e.g. for DDLC's poems and fake console:
  `--extra "(title|text)\s*=\s*$" --extra "updateconsole\s*\(.*,\s*$"`.
* Re-running `extract` on an existing workbook keeps the translations of rows
  whose original text did not change.

After extracting, the script asks what to do next (translate with Gemini, build
the context only, or nothing). Use `--no-prompt` to skip the question or
`--translate` to start translating right away.

## 2. Translate with Gemini

```
python vntextpatch.py translate script.xlsx --keys-file keys.txt
```

* **Keys**: `keys.txt` has one API key per line (or list them in the config
  file, or `--keys k1,k2`, or the `GEMINI_API_KEYS` environment variable). Requests rotate round-robin; a key
  that returns 429 rests for the delay the API asks for and the next key takes
  over; invalid keys are dropped.
* **Context ("memory")**: the first run shows Gemini a sample of the script and
  saves `script.context.json`: a glossary with one fixed Burmese spelling per
  name/term, a character sheet (gender, personality, speech style, how they
  address each other), story notes and a style guide. Edit that file freely
  (there is also an `"extra_notes"` field for your own instructions); names you
  translate by hand in `names.xml` override it.
* **Context caching**: the rules + that context are uploaded once per key as a
  Gemini `cachedContent` and referenced by every request. If a key or model
  refuses to create the cache (content below the model's minimum size, tier
  without caching) the identical prefix is sent inline instead, which Gemini's
  implicit caching still discounts. Cache ids are kept in
  `script.gemini_cache.json` (keys are stored only as hashes); the cache is
  rebuilt when the context changes. `--delete-cache` removes them at the end.
* Each request also carries the previous 8 lines and their translations.
* Translations whose `[placeholders]`, `{tags}` or `\n` differ from the source
  are retried and, if still different, flagged with `CHECK:` in `Notes`.
* The run is resumable: rows that already have a `TL` value are skipped.

* **Batch size**: `--batch-size N` sets how many lines go into one request
  (default 40) and `--batch-chars N` caps the source characters per request
  (default 6000); a batch closes at whichever limit is reached first. In the
  menu both are under Settings > Batch size. Larger batches use fewer requests,
  which helps with request-per-day quotas; smaller ones are cheaper to retry.

* **Progress**: a progress bar shows lines done, speed, ETA and the current
  sheet. The workbook is saved after every batch, so the translated text is in
  the `.xlsx` as it arrives and an interrupted run loses nothing
  (`--save-every N` saves every N batches instead).
* **Errors**: every error Gemini returns is printed with the model, the key
  (last 4 characters only), the HTTP status and message, and what the script
  does about it. The full responses are appended to
  `script.gemini_errors.log`.

### Config file and fallback models

Gemini settings live in a dedicated `gemini_config.json`. Create one with
`python vntextpatch.py config` (the menu creates it automatically); `translate`
looks for it next to the workbook, then in the current folder, or takes
`--config PATH`. Command line options override it.

```json
{
  "models": ["gemini-2.5-flash", "gemini-2.5-flash-lite", "gemini-2.0-flash"],
  "keys_file": "keys.txt",
  "keys": [],
  "batch_size": 40,
  "batch_chars": 6000,
  "workers": 0,
  "temperature": 0.3,
  "thinking_budget": null,
  "cache": true,
  "cache_ttl": 3600,
  "max_attempts": 12,
  "fallback_after_errors": 3,
  "model_rest_seconds": 60,
  "request_timeout": 180,
  "sample_chars": 120000,
  "error_log": "",
  "api_base": "https://generativelanguage.googleapis.com/v1beta"
}
```

`models` is the main model followed by its fallbacks, in order. The script
always uses the first model that is available and moves down the list when:

* every key has hit its quota on that model (quotas are tracked per key *and*
  per model, so the same keys keep working on the fallback),
* the model answers with server errors or broken JSON
  `fallback_after_errors` times in a row (it is then rested for
  `model_rest_seconds`),
* the model does not exist or is not supported (it is dropped for the run).

It returns to the main model as soon as that is available again. The summary
at the end shows how many requests each model served. Each key/model pair gets
its own context cache.

Other options: `--model a,b,c` (main model + fallbacks), `--language`,
`--sheets a,b`, `--limit 20` (test run), `--workers`, `--context-only`,
`--rebuild-context`, `--retranslate`, `--no-cache`, `--cache-ttl`.

## 3. Insert

```
python vntextpatch.py insert <same input as extract> script.xlsx <output folder> [--repack]
```

For each row the text used is `Edit`, else `TLC`, else `TL`, else the original
(`.` skips a column, `(empty)` means an empty string) - as in VNTextPatch.
Without `--repack` the patched `.rpy` files are written to the output folder;
with it, each input archive is rebuilt there with the patched scripts (stale
`.rpyc` of patched scripts are left out so Ren'Py recompiles them).
Pass the same `--mode` / `--extra` options that were used for `extract`.
