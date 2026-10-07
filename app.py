# Declaration: Code generated using Anthropic Claude (Opus 5.5)
import json
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

import proofreader as pr

OUTPUT_ROOT = "output"
SETTINGS_FILE = Path(__file__).parent / "settings.json"
AUTOMATIC = "Automatic"
CONTEXT_CHOICES = [AUTOMATIC] + pr.CONTEXT_OPTIONS
DISPLAY_INTERVAL_SECONDS = 0.5

DEFAULTS = {
    "model": "",
    "target_words": pr.TARGET_WORDS,
    "num_ctx": AUTOMATIC,
    "page_offset": 0,
    "footnote_ratio": pr.FOOTNOTE_SIZE_RATIO,
    "think": False,
    "temperature": pr.TEMPERATURE,
    "include_notes": True,
    "output_root": OUTPUT_ROOT,
    "prompt": pr.DEFAULT_PROMPT,
    "language": "English (USA)",
}

PAGE_STYLE = """
<style>
h1, h2, h3 { font-family: Georgia, "Times New Roman", serif; font-weight: 600; }
.block-container { padding-top: 2.5rem; }
</style>
"""


def format_seconds(seconds):
    seconds = int(seconds)
    hours, remainder = divmod(seconds, 3600)
    minutes, secs = divmod(remainder, 60)
    if hours:
        return f"{hours}h {minutes:02d}m"
    if minutes:
        return f"{minutes}m {secs:02d}s"
    return f"{secs}s"


def load_settings():
    try:
        return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def save_settings(settings):
    try:
        SETTINGS_FILE.write_text(json.dumps(settings, indent=2), encoding="utf-8")
    except OSError:
        pass


def reset_prompt():
    st.session_state["prompt"] = pr.DEFAULT_PROMPT


@st.cache_data(show_spinner="Converting the Word document to PDF...")
def convert_upload(data, filename):
    return pr.convert_word_to_pdf(data, filename)


@st.cache_data(show_spinner="Reading PDF...")
def read_pdf(pdf_bytes, page_offset, footnote_ratio):
    return pr.extract_document(pdf_bytes, page_offset, footnote_ratio)


@st.cache_data(ttl=300, show_spinner=False)
def cached_model_info(model):
    return pr.get_model_info(model)


def show_progress(state, widgets, stage="", reply="", thinking=""):
    now = time.time()
    chunk_elapsed = now - state["chunk_started"]
    todo_words = state["todo_words"]
    fraction = state["done_words"] / todo_words
    remaining_text = "estimating..."
    seconds_per_word = None
    if state["words_done_run"]:
        seconds_per_word = state["seconds_spent"] / state["words_done_run"]
    elif state["stored_rate"]:
        seconds_per_word = state["stored_rate"]
    if seconds_per_word:
        expected_chunk = seconds_per_word * state["chunk_words"]
        partial = min(chunk_elapsed / expected_chunk, 0.95) if expected_chunk > 0 else 0
        fraction = (state["done_words"] + partial * state["chunk_words"]) / todo_words
        current_left = max(expected_chunk - chunk_elapsed, 0.1 * expected_chunk)
        later = seconds_per_word * (todo_words - state["done_words"] - state["chunk_words"])
        remaining_text = format_seconds(current_left + later)
    fraction = min(fraction, 1.0)

    activity = "thinking..." if thinking and not reply else "writing..."
    widgets["bar"].progress(fraction, text=f"Chunk {state['position']} of {state['total']} · {int(fraction * 100)}%")
    widgets["chunk"].metric("Chunk", f"{state['position']} of {state['total']}")
    widgets["elapsed"].metric("Elapsed", format_seconds(now - state["started"]))
    widgets["remaining"].metric("Remaining", remaining_text)
    stage_text = f"{stage}, {activity}" if stage else "starting..."
    widgets["status"].write(f"{state['pages']} · {stage_text} {format_seconds(chunk_elapsed)} on this chunk")


def update_display(state, widgets, stage, reply, thinking):
    now = time.time()
    if now - state["last_draw"] >= DISPLAY_INTERVAL_SECONDS:
        state["last_draw"] = now
        show_progress(state, widgets, stage, reply, thinking)


st.set_page_config(page_title="Proof Reader", layout="centered")
st.markdown(PAGE_STYLE, unsafe_allow_html=True)

saved_settings = load_settings()
for key, default in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = saved_settings.get(key, default)
if st.session_state["num_ctx"] not in CONTEXT_CHOICES:
    st.session_state["num_ctx"] = AUTOMATIC

with st.sidebar:
    st.header("Settings")
    with st.expander("Model", expanded=True):
        if "models" not in st.session_state or st.button("Refresh model list"):
            try:
                st.session_state["models"] = pr.list_models()
                st.session_state["models_error"] = ""
            except Exception as error:
                st.session_state["models"] = []
                st.session_state["models_error"] = str(error)
        models = st.session_state["models"]
        if st.session_state["models_error"]:
            st.error("Could not reach Ollama. Is it running? " + st.session_state["models_error"])
        elif not models:
            st.warning("No models found. Use 'ollama pull <model>' in a terminal.")
        model = None
        model_info = None
        if models:
            if st.session_state["model"] not in models:
                st.session_state["model"] = models[0]
            model = st.selectbox("Model", models, key="model")
            model_info = cached_model_info(model)
            st.caption(pr.describe_model(model_info))
        think_requested = st.checkbox(
            "Let thinking models think (slower)",
            key="think",
            help="Thinking can improve results but is much slower. The reasoning is saved in each chunk file.",
        )
        think = False
        if think_requested and model_info is not None:
            if model_info["supports_thinking"]:
                think = True
                st.caption(
                    "Thinking is on. Expect much longer runs. The context size is raised automatically, "
                    "and if the model thinks for too long the chunk is redone without thinking."
                )
            else:
                st.warning("This model does not report thinking support, so it will run without it.")
        temperature = st.slider(
            "Temperature", 0.0, 1.0, step=0.05, key="temperature",
            help=f"Lower gives more consistent, literal answers; higher gives more varied ones. Default {pr.TEMPERATURE}.",
        )
    with st.expander("Chunking", expanded=True):
        include_notes = st.toggle(
            "Include footnotes", key="include_notes",
            help="On: footnotes are proofread in a second pass and count towards the chunk size. "
            "Off: footnotes are left out entirely and only the main text is proofread.",
        )
        target_words = st.number_input(
            "Chunk size (words, including footnotes)" if include_notes else "Chunk size (words)", min_value=300, max_value=3000, step=50, key="target_words",
        )
        page_offset = st.number_input(
            "Page offset", step=1, key="page_offset",
            help="Added to the PDF page number, e.g. -4 if page 1 of the text is PDF page 5.",
        )
        footnote_ratio = st.slider(
            "Footnote font size (fraction of body size)", 0.5, 1.0, step=0.01, key="footnote_ratio",
            help="Text smaller than this, below the main text on a page, is treated as footnotes.",
        )
    with st.expander("Advanced"):
        context_choice = st.selectbox("Context size (tokens)", CONTEXT_CHOICES, key="num_ctx")
        output_root = st.text_input("Output folder", key="output_root")

st.title("Proof Reader")
st.caption("Proofreads a PDF or Word document in chunks with a local model through Ollama. Nothing leaves your computer.")

chunks = []
source_name = ""
converted_by = None
num_ctx = pr.CONTEXT_OPTIONS[0]

with st.container(border=True):
    st.subheader("1 · Document")
    uploaded = st.file_uploader(
        "PDF or Word document to proofread", type=["pdf", "docx", "doc"], label_visibility="collapsed",
    )
    if uploaded is not None:
        source_name = Path(uploaded.name).stem
        pdf_bytes = uploaded.getvalue()
        if pr.is_word_file(uploaded.name):
            try:
                pdf_bytes, converted_by = convert_upload(pdf_bytes, uploaded.name)
                note = " Page numbers match Word's layout." if converted_by == "Microsoft Word" else (
                    " Page numbers may differ slightly from Word's layout."
                )
                st.caption(f"Converted to PDF with {converted_by}.{note}")
            except RuntimeError as error:
                st.error(str(error))
                pdf_bytes = None
        paragraphs, info = read_pdf(pdf_bytes, int(page_offset), footnote_ratio) if pdf_bytes else ([], {})
        if pdf_bytes and not paragraphs:
            st.error("No text found. The PDF may be a scan without a text layer.")
        elif paragraphs:
            chunks = pr.make_chunks(paragraphs, int(target_words), include_notes)
            current_prompt = pr.full_prompt(st.session_state["prompt"], st.session_state["language"])
            if context_choice == AUTOMATIC:
                num_ctx = pr.choose_context(chunks, current_prompt, think)
            else:
                num_ctx = int(context_choice)
            needed = max(pr.context_needed(chunk, current_prompt, think) for chunk in chunks)
            if needed > num_ctx:
                st.warning(
                    f"The largest chunk may need about {needed:,} tokens, more than the context size of "
                    f"{num_ctx:,}. Ollama would silently cut the text off. Choose a larger context size "
                    "in Advanced, or a smaller chunk size."
                )
            if info["unlinked"] > max(3, info["notes"] * 0.05):
                st.warning(
                    f"{info['unlinked']} of {info['notes']} footnotes could not be matched to a reference number "
                    "in the text. They were placed with the last paragraph on their page."
                )
            if info["repaired"]:
                fixes = ", ".join(f"{change} ({count}×)" for change, count in info["repaired"].items())
                st.info(
                    "This PDF stores some letters wrongly (usually ligatures such as “ff” recorded as another "
                    f"character, which happens with the Aptos font). Repaired: {fixes}. Check the preview."
                )
            st.caption(
                f"{len(chunks)} chunks · {info['paragraphs']} paragraphs · {info['notes']} footnotes · "
                f"context size {num_ctx:,} tokens{' (automatic)' if context_choice == AUTOMATIC else ''}"
            )
            italics = pr.italics_consistency(chunks)
            with st.expander(f"Italics consistency ({len(italics)} terms to check)"):
                st.caption(
                    "Found by the app without the model: titles and non-English words that are italic in some "
                    "places but not in others. This list is also included in the full report."
                )
                st.markdown("\n".join(italics) if italics else "No inconsistencies found.")
            with st.expander("Preview chunks"):
                columns = st.columns(4)
                columns[0].metric("Chunks", len(chunks))
                columns[1].metric("Main text words", f"{sum(c['body_words'] for c in chunks):,}")
                columns[2].metric("Footnote words", f"{sum(c['note_words'] for c in chunks):,}")
                columns[3].metric("Footnotes", info["notes"])
                table = pd.DataFrame([
                    {
                        "Chunk": chunk["number"],
                        "Pages": pr.page_range_label(chunk),
                        "Total words": chunk["words"],
                        "Main text": chunk["body_words"],
                        "Footnotes": chunk["note_words"],
                        "Starts": chunk["main_text"][:70].replace("\n", " "),
                    }
                    for chunk in chunks
                ])
                st.dataframe(table, hide_index=True)
                chosen = st.selectbox("Show the text sent to the model for chunk", [c["number"] for c in chunks])
                st.markdown("**Main text**")
                st.text(chunks[chosen - 1]["main_text"])
                if chunks[chosen - 1]["notes_text"]:
                    st.markdown("**Footnotes**")
                    st.text(chunks[chosen - 1]["notes_text"])

with st.container(border=True):
    st.subheader("2 · Prompt")
    language = st.text_input(
        "Language", key="language",
        help="The language and spelling conventions to check against, e.g. English (USA), English (UK), French.",
    )
    prompt = st.text_area("Prompt", key="prompt", height=140)
    model_prompt = pr.full_prompt(prompt, language)
    st.button("Reset prompt to default", on_click=reset_prompt)
    with st.expander("Instructions added automatically after your prompt"):
        st.markdown("**For the main text**")
        st.code(pr.OUTPUT_INSTRUCTIONS, language=None, wrap_lines=True)
        st.markdown("**For the footnotes**")
        st.code(pr.FOOTNOTE_INSTRUCTIONS, language=None, wrap_lines=True)

current_settings = {key: st.session_state[key] for key in DEFAULTS}
if current_settings != saved_settings:
    save_settings(current_settings)

if uploaded is None:
    st.info("Choose a PDF or Word document to begin.")
    st.stop()
if not chunks:
    st.stop()

with st.container(border=True):
    st.subheader("3 · Proofread")
    existing = pr.existing_reports(Path(output_root) / source_name)
    mode = pr.RESUME
    start_label = "Start proofreading"
    if existing:
        saved_at = datetime.fromtimestamp(max(path.stat().st_mtime for path in existing))
        st.info(
            f"Earlier output found for “{source_name}”: {len(existing)} chunk reports, "
            f"last saved {saved_at:%d %b %Y, %H:%M}."
        )
        mode = st.radio(
            "What would you like to do?",
            [pr.RESUME, pr.OVERWRITE, pr.KEEP_BOTH],
            captions=[
                "Keep finished chunks (only chunks with the same model, prompt and settings count as finished).",
                "Delete the earlier chunk reports and full report for this file, and start again.",
                "Leave the earlier output alone and write to a new folder.",
            ],
        )
        start_label = {
            pr.RESUME: "Resume proofreading",
            pr.OVERWRITE: "Overwrite and start",
            pr.KEEP_BOTH: "Start as new copy",
        }[mode]

    button_columns = st.columns([1, 1, 2])
    start_clicked = button_columns[0].button(start_label, type="primary", disabled=model is None)
    button_columns[1].button("Stop", help="Finished chunks are kept. Press Start and choose Resume to continue.")

    if start_clicked:
        out_dir = pr.choose_output_folder(output_root, source_name, mode)
        model_text = pr.describe_model(model_info)
        rate_key = pr.timing_key(model, think)
        stored_rate = pr.load_timings(output_root).get(rate_key)

        jobs = []
        for chunk in chunks:
            path = out_dir / pr.chunk_filename(chunk)
            signature = pr.chunk_signature(chunk, model, model_prompt, num_ctx, think, temperature)
            finished = pr.load_saved_report(path, signature) is not None
            jobs.append({"chunk": chunk, "path": path, "signature": signature, "finished": finished})
        todo = [job for job in jobs if not job["finished"]]
        todo_words = sum(job["chunk"]["words"] for job in todo)
        skipped = len(jobs) - len(todo)

        bar = st.progress(0.0, text="Starting...")
        metric_columns = st.columns(3)
        widgets = {
            "bar": bar,
            "chunk": metric_columns[0].empty(),
            "elapsed": metric_columns[1].empty(),
            "remaining": metric_columns[2].empty(),
            "status": st.empty(),
        }
        latest = st.empty()
        if skipped:
            st.caption(f"{skipped} chunk(s) already finished were kept.")

        started = time.time()
        state = {
            "started": started, "total": len(todo), "todo_words": todo_words, "done_words": 0,
            "words_done_run": 0, "seconds_spent": 0.0, "stored_rate": stored_rate,
            "position": 0, "chunk_words": 0, "chunk_started": started, "pages": "", "last_draw": 0.0,
        }

        for position, job in enumerate(todo, start=1):
            chunk = job["chunk"]
            state.update(
                position=position, chunk_words=chunk["words"], chunk_started=time.time(),
                pages=pr.page_range_label(chunk), last_draw=0.0,
            )
            show_progress(state, widgets)

            result = None
            failure = ""
            for attempt in (1, 2):
                try:
                    result = pr.proofread_chunk(
                        chunk, model, model_prompt, num_ctx, think,
                        lambda stage, reply, thought: update_display(state, widgets, stage, reply, thought),
                        temperature,
                    )
                    break
                except Exception as error:
                    failure = str(error)
            if result is None:
                st.error(
                    f"Chunk {chunk['number']} ({pr.page_range_label(chunk)}) failed twice: {failure}\n\n"
                    f"Finished chunks are saved in {out_dir.resolve()}. "
                    "Fix the problem, then press Start and choose Resume to continue."
                )
                st.stop()

            pr.save_chunk_report(
                job["path"], chunk, len(chunks), model_text, model_prompt, num_ctx, think, job["signature"], result,
                temperature,
            )
            seconds = time.time() - state["chunk_started"]
            pr.update_timing(output_root, rate_key, seconds, chunk["words"])
            state["seconds_spent"] += seconds
            state["words_done_run"] += chunk["words"]
            state["done_words"] += chunk["words"]
            latest.text(
                f"Latest: chunk {chunk['number']} ({pr.page_range_label(chunk)})\n\n{pr.format_chunk_body(result)}"
            )

        full_report = pr.build_full_report(
            chunks, out_dir, uploaded.name, model_text, think, prompt, int(target_words), num_ctx,
            converted_by, include_notes, language, temperature,
        )
        bar.progress(1.0, text="Finished")
        total_seconds = time.time() - started
        widgets["status"].success(
            f"Finished in {format_seconds(total_seconds)}. Files saved in {out_dir.resolve()}"
        )
        st.session_state["result"] = {"source": source_name, "full_report": full_report}

    result = st.session_state.get("result")
    if result and result["source"] == source_name:
        st.divider()
        st.subheader("Full report")
        report_name = pr.full_report_filename(source_name)
        st.download_button(
            f"Download {report_name}", result["full_report"], file_name=report_name, mime="text/plain",
        )
