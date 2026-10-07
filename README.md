# Test Bank Audit Pipeline — Rebuild

Full rebuild of the Cengage test bank audit pipeline. All core logic lives in
a single `src/testbank_pipeline.py` module, organised into four clearly
separated sections. The Streamlit UI (`app.py`) sits at the project root and
delegates everything to the pipeline class.

## Project structure

```
testbank_pipeline_rebuild/
├── app.py                      # Streamlit UI
├── requirements.txt
├── .env / .example.env         # Gemini API key & tunables
├── prompts/
│   └── testbank.txt            # Review prompt
├── src/
│   └── testbank_pipeline.py    # All pipeline logic (see sections below)
├── logs/
│   ├── cost_logger.py          # Cost tracking & Excel logging
│   ├── output_logger.py        # Per-item decision logging
│   ├── calculate_logs.py       # Summary sheet refresh
│   ├── cost_log.xlsx           # Accumulated cost data
│   └── output_log.xlsx         # Accumulated output decisions
```

## `src/testbank_pipeline.py` — section map

All pipeline logic is consolidated into one file, organised in four sections:

| Section | Contents |
|---|---|
| **Section 1 — Shared Data Models** | `TestBankItem`, `OutputItem`, `ChunkStats` dataclasses; field/label constants (`METADATA_FIELDS`, `METADATA_LABELS`, `OPTION_LETTERS`); decision constants (`KEEP`, `REVISE`, `REMOVE`, `PARSE_ERROR`). |
| **Section 2 — DOCX I/O** | `.doc → .docx` conversion (LibreOffice / Word COM); PDF text extraction; `ChapterExtractor`; `TestBankParser`; `DocxBuilder` with tracked-change run formatting. |
| **Section 3 — Gemini Engine** | `GeminiClient` (google-genai SDK wrapper); `PromptBuilder`; `ResponseParser`; `GeminiChunkedProcessor`; `ContentValidator` with Cengage guideline rules. |
| **Section 4 — Pipeline Orchestrator** | `TestBankPipeline` class that wires all sections together; `PipelineResult` dataclass; cost/output logging integration; CLI `main()` entry point. |

## Why this design is DRY / OOP

- **One data model.** `TestBankItem` and `OutputItem` are defined once, including the metadata field list and label text — the parser, prompt builder, response parser, and docx builder all use the same constants instead of re-typing `"ANSWER:\xa0\xa0"` in multiple places.
- **One conversion layer.** `.doc` conversion and PDF reading are defined once and used by both `TestBankParser` and `ChapterExtractor`.
- **One rule engine, not scattered `if` statements.** `ContentValidator` registers each Cengage guideline as a small method in a list; adding a new rule is a one-line addition, not a new branch buried in the docx builder.
- **The orchestrator has almost no logic.** `TestBankPipeline.process_test_bank` just calls each collaborator in order — every actual decision (how to parse, how to chunk, how to validate, how to build the docx) lives in the class that owns that concern, so each piece can be unit-tested without booting the whole pipeline or hitting the API.
- **Guaranteed no silent drops.** `ResponseParser` always emits one `OutputItem` per input `TestBankItem`; a parse failure becomes a `PARSE_ERROR` item carrying the original data forward, never a gap.

## Environment variables

See `.example.env` for the full list. Key variables:

```
GEMINI_API_KEY=your_key     # Required
GEMINI_MODEL_ID=gemini-3.5-flash-lite  # Model to use
GEMINI_TEMPERATURE=0.1      # 0.0 = deterministic
MAX_OUTPUT_TOKENS=65536     # Max response tokens
CHUNK_SIZE=10               # Items per Gemini call
MAX_RETRIES=2               # Chunk retry limit
```

## Running

```bash
pip install -r requirements.txt --break-system-packages
streamlit run app.py
```
