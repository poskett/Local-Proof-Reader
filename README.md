# Proof Reader

Proofreads a long PDF or Word document with a local model through Ollama. Nothing leaves your computer.

The app rebuilds the PDF's paragraphs and footnotes from the page layout. It splits the text into chunks of about 1,500 words including footnotes, never mid-paragraph and never mid-sentence unless a single paragraph is longer than the limit. Each chunk is proofread in two passes, main text then footnotes. Python then finds every quoted error in the source to supply its page (and footnote) number, and stitches the chunk reports into one full report.

## Set-up

1. Install [Ollama](https://ollama.com) and pull a model, e.g. `ollama pull qwen3-vl:8b-instruct-q4_K_M`.
2. `pip install -r requirements.txt`
3. `streamlit run app.py`

To stop a long run being interrupted by the Mac going to sleep, start the app with `caffeinate -i streamlit run app.py`. Closing the browser tab also stops a run (finished chunks are kept).

## Use

1. Choose a model in the sidebar (press *Refresh model list* after pulling a new one).
2. Upload a PDF or Word document (.docx or .doc). The line under the upload box gives the number of chunks, paragraphs and footnotes found. Open *Preview chunks* to see exactly what the model will be sent.
3. Set the language (default *English (USA)*; type any other, e.g. *English (UK)*). Edit the prompt if you wish, then press *Start proofreading*.
4. Watch the progress bar, elapsed time and time remaining. The first estimate uses the speed recorded from earlier runs with the same model. After the first chunk it uses this run's speed.
5. *Stop* interrupts the run. Finished chunks stay saved.

### If the file has been proofread before

If `output/<pdf name>/` already has chunk reports, you choose:

- **Resume**: keep finished chunks. A chunk counts as finished only if the model, thinking setting, temperature, prompt, context size and chunk text are all unchanged.
- **Overwrite**: delete the earlier chunk reports and full report for this file, and start again. Other files in the folder are left alone.
- **Keep both**: write to a new folder (`<pdf name>_2`, `_3`, ...).

## Word documents

A Word document is converted to PDF automatically when you upload it. The converted PDF is not saved.

- **With Microsoft Word installed (Mac),** the app asks Word to save a PDF copy in the background. Page numbers then match what you see in Word, and Word's own PDF export avoids the Aptos "ff → X" problem described below. The first time, macOS may ask whether to allow the app (or Terminal) to control Microsoft Word: click OK. If Word wasn't already open, it is closed again afterwards. Your open documents are left untouched, because Word works on a temporary copy.
- **Otherwise, LibreOffice** is used if it is installed. The text, footnotes and italics come out the same, but LibreOffice substitutes some fonts, so page numbers can drift by a page or two from Word's. The app says which converter was used, and the full report records it.

## How the text is prepared

- **Paragraphs** are rebuilt from the layout (first-line indents, spacing, short last lines, headings), not from how the PDF happens to store its text. This matters for PDFs made with *Print → Save as PDF* on a Mac, which store every line separately. A paragraph running onto the next page is kept whole.
- **Footnote reference numbers** (superscripts) are removed from the main text the model sees. The app still uses them to keep each footnote with the paragraph that cites it.
- **Footnotes** are the small text below the main text on each page. A note continuing onto the next page is joined up. Each chunk's footnotes are sent in a second pass with instructions suited to citations.
- **Ligatures** stored as special characters (ﬁ, ﬀ) are expanded to ordinary letters. Some PDFs store a ligature as the wrong letter altogether. If you can, export with Word's *File → Save As → PDF*, which usually avoids the problem.
- **Line ends** ending in a hyphen or dash are joined without a space, keeping the hyphen ("semi-nomadic", "21-4").

## Italics

The model is sent plain text, without italics. In testing, marking italics for the model confused it more than it helped, and the default prompt tells it not to check them. Instead, the app checks italics itself:

- **Consistency across the whole document.** It lists titles and non-English words that are italic in some places but roman (or partly italic) in others, with page and footnote numbers. The list appears straight away under the upload box and as a section in the full report.
- To avoid false alarms, it ignores roman occurrences inside quotation marks (article and chapter titles). 
- **Limits:** a title that is never italicised anywhere can't be caught. A title within an italic title, correctly set in roman, will appear in the list. Check it and move on.

## Output

Saved in `output/<pdf name>/`:

- `chunk_01_pp1-5.txt`, ...: one report per chunk (plain text), recording the model, thinking setting, temperature and prompt, with *Main text* and *Footnotes* sections.
- `<pdf name>_full_report.txt`: a plain-text file that opens in any text editor or in Word. It contains the model name and details (size, quantisation, digest), language, prompt and settings; a contents table with the number of items flagged per chunk; the italics consistency check; then every chunk report in order. The app's *Download* button gives you the same file.

## Checks built in

- **Cut-off checks:** a chunk fails (and is retried once) if the reply hit its length limit, or if the text filled the context window. If it fails twice the run stops with a plain message and finished chunks are kept.
- **Thinking models:** off by default. If you tick it, the app checks the model supports thinking and raises the context size. Thinking is capped at about 6,000 tokens (a few minutes per pass). If the model is still thinking at that point, the pass is redone without thinking, and the report says so. In testing, gemma4 often thought far longer than this on 1,500-word chunks, so thinking may add time without changing the result.
- Repeated identical lines in a report are collapsed to one line with a count, e.g. `(×3)`.

## Settings (sidebar)

Settings, the language and the prompt are remembered in `settings.json` next to `app.py`. 

- Include footnotes: on by default. Switch off to leave footnotes out entirely.
- Chunk size: target words per chunk, including footnotes when they are included (the limit is 20% above the target).
- Page offset: added to PDF page numbers (e.g. -4 if the text's page 1 is PDF page 5).
- Footnote font size: text smaller than this fraction of the body size, below the main text, is treated as footnotes.
- Temperature (Model section): 0.4 by default. Lower values give more consistent, literal answers; higher values more varied ones. It is recorded in each report.
- Context size: *Automatic* (default) picks the smallest size that fits the largest chunk plus room for the reply: usually 8,192 tokens, or 12,288 with thinking. One size is used for the whole run so the model isn't reloaded between chunks.
- Output folder.

## Limitations

- Scanned PDFs without a text layer are not supported.
- The layout rules assume a single column of text. Endnotes gathered at the end of a document are treated as ordinary text.
- Page numbers are PDF page numbers plus the offset.
- Local models miss errors, and flag correct historical spellings, quotations and names. Tell the model your conventions in the prompt (e.g. British spelling, how page ranges are written) to reduce false "consistency" suggestions. Treat every suggestion as something to check.
- As the model receives the text in chunks, it cannot check for consistency if the errors are many pages apart (ie. a consistency problem between p.2 and p.56). Increase the chunk size to address this if your hardware allows (at the expense of accuracy).
