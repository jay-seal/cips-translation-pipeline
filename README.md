# CIPS Translation Pipeline — PPTX Reconstruction

This repository contains the reconstruction step for the CIPS translation pipeline.
It takes the JSON outputs from the Opal translation agents and applies them to the
original PPTX, producing a translated PPTX file for download.

---

## How to use

### Step 1 — Prepare your input files

You need three files from the Opal pipeline:

| File | Where it comes from | Required |
|---|---|---|
| `source.pptx` | The original PPTX you are translating | Yes |
| `agent3_output.json` | The JSON export from Agent 3 (Translation Agent) | Yes |
| `agent4_output.json` | The JSON export from Agent 4 (QA Agent) | No — but recommended |

Rename your source PPTX to `source.pptx` before uploading.

### Step 2 — Upload files to the inputs/ folder

In this GitHub repository, navigate to the `inputs/` folder and upload your three files using the **Add file → Upload files** button in the GitHub interface.

Commit the upload directly to the main branch.

### Step 3 — The workflow runs automatically

As soon as files land in `inputs/`, the GitHub Actions workflow fires automatically.
You can watch progress under the **Actions** tab.

The workflow takes approximately 30–60 seconds to complete.

### Step 4 — Download the translated PPTX

When the workflow completes:
1. Click the workflow run in the **Actions** tab
2. Scroll to the **Artifacts** section at the bottom of the page
3. Click the artifact name to download the translated PPTX

The output file is named automatically: `source_FR-FR.pptx` (or whatever locale
was used in the translation).

Artifacts are retained for **30 days**.

---

## Running manually

If you want to trigger a run without uploading new files (e.g. to re-run after
fixing a JSON file already in the inputs/ folder):

1. Go to the **Actions** tab
2. Select **CIPS Translation — Reconstruct PPTX**
3. Click **Run workflow → Run workflow**

---

## What the script does

- Matches each translated segment to its shape on the correct slide using
  slide number + source text
- Replaces text while preserving run-level formatting (font, size, bold,
  italic, colour)
- Skips KEPT and IMAGE_TEXT segments automatically
- Applies CRITICAL capitalisation and punctuation corrections from the QA Agent
- Reduces font size by up to 2pt on slides with CRITICAL expansion issues
- Produces a reconstruction report in the workflow log

---

## Adding a new locale

To translate into a different language:

1. Run the Opal pipeline (Agents 1–4) with the new `locale_code`
2. Export the Agent 3 and Agent 4 JSON outputs
3. Upload to `inputs/` as above — the output filename will reflect the locale
   automatically

No changes to the script or workflow are needed for new locales.

---

## Troubleshooting

**The workflow didn't trigger after upload**
Make sure you committed the files to the main branch, not a different branch.

**Some segments are unmatched**
Check the workflow log for the reconstruction report. Unmatched segments usually
mean the text in the PPTX shape doesn't exactly match what was extracted during
ingestion — this can happen when text is split across multiple runs with different
formatting. These slides will retain their original English text and are listed in
the report for manual review.

**The QA Agent set routing to SME_REVIEW_REQUIRED**
The script will block and exit if the Agent 4 output contains
`routing: "SME_REVIEW_REQUIRED"`. Resolve the CRITICAL issues in the Opal Review
Canvas before running reconstruction.
