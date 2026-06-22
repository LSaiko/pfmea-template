# LLM Cold-Start Feature — Requirements & Integration Design

**Feature:** `--llm-generate` — optional, fully-offline local-LLM scenario drafting.
**Status:** Design / pre-implementation (NOT yet built — this is the spec).

> **Integration note for this repo's actual layout.** The design below references
> a `pfmea/` package (`generator.py`, `cli.py`, `schema.py`). The current code is
> flat: [`pfmea.py`](../pfmea.py) is the engine + CLI, scenarios are YAML files in
> [`scenarios/`](../scenarios), and `export.py` renders Word/PDF. When implemented,
> the LLM code goes in a single **`llm.py`** module, lazy-imported only when
> `--llm-generate` is passed, exposing `draft_scenario(...) -> dict`. The drafted
> YAML is written to `scenarios/` and then re-enters the normal
> `pfmea.py generate` pipeline unchanged. `get_action_priority` stays the single
> source of truth for AP — the LLM never writes AP.

---

## 1. Goals & Non-Goals

**Goals**
- Bootstrap a scenario YAML from a free-text process description in <60 s, fully offline.
- Output is a *draft* conforming to the scenario schema; it enters the render pipeline unchanged.
- Every LLM field tagged `draft: true`; S/O/D flagged `DRAFT — REQUIRES ENGINEER REVIEW`
  (research shows LLMs are unreliable at numeric risk ratings — see §9).
- AP is never written by the LLM; always computed by `get_action_priority(s,o,d)`.
- Full prompt + raw response logged to a sidecar `.audit.log` for 21 CFR 820 / ISO 13485 traceability.

**Non-Goals**
- No finalize/approve/publish; no cloud calls; no telemetry; no network egress.
- Does not replace human review — drafting accelerator only.
- Base tool stays installable with zero new *required* dependencies.

## 2. Dependencies & Runtime Options

Optional extras only (never pulled in by default): `httpx`, `jsonschema` (PyYAML already present).

| Runtime | Default endpoint | Notes |
|---|---|---|
| Ollama | `http://localhost:11434/api/generate` | supports `format: json` |
| LM Studio / llama.cpp | `http://localhost:1234/v1/chat/completions` | OpenAI-compatible |
| Any OpenAI-compatible local server | configurable | same client path |

Recommended models (8–16 GB VRAM): `llama3.1:8b`, `mistral:7b-instruct`, `phi3:mini`.
Auto-detect API flavor from endpoint path (`/api/generate` → Ollama; `/v1/chat/completions` → OpenAI-compat).

## 3. CLI Surface

```bash
python pfmea.py generate \
  --llm-generate \
  --describe "heat-sealing Tyvek pouches for sterile orthopedic implants" \
  --model llama3.1:8b \
  --endpoint http://localhost:11434 \
  -o scenarios/tyvek_heatseal_draft.yaml \
  --rag        # optional grounding from scenarios/
```

| Flag | Default | Description |
|---|---|---|
| `--llm-generate` | off | enable; errors if `llm` extras absent |
| `--describe TEXT` | required | free-text process description |
| `--model` | `llama3.1:8b` | model name |
| `--endpoint` | `http://localhost:11434` | local runtime base URL |
| `--rag` | off | retrieval-augmented grounding from `scenarios/` |
| `--force` | off | allow overwriting an existing scenario |

Config precedence: `~/.pfmea.toml` → env (`PFMEA_LLM_ENDPOINT`, `PFMEA_LLM_MODEL`) → flags.

## 4. Prompt Design

Single-turn structured prompt with schema injection + few-shot examples; constrained
output via Ollama `format: json` (or `response_format: {type:"json_object"}`). System
prompt states: medical-device pFMEA assistant under ISO 13485; return ONLY valid JSON
matching the schema; S/O/D integers 1–10; mark every value `draft: true`. Few-shot
examples pulled at runtime from the smallest 1–2 files in `scenarios/` (so the model
matches this project's exact schema). Temperature 0.2 for determinism.

## 5. Output Validation & Safety

1. JSON parse (log raw on failure). 2. `jsonschema.validate` against the shared scenario
schema. 3. Clamp S/O/D to [1,10], log warnings. 4. Inject `_llm_draft: true`, `_llm_model`,
`_llm_timestamp`. 5. Overwrite guard (abort unless `--force`). 6. Audit sidecar
`<out>.audit.log` with timestamp, model, endpoint, full prompts, raw response, warnings.
The draft goes straight to the human engineer — never auto-revised by the LLM.

## 6. RAG Option (stretch)

With `--rag`: load `scenarios/`, TF-IDF (or token-overlap fallback) similarity between
`--describe` and each scenario's process description, inject top-2 as few-shot examples.
Grounding on domain-matched, previously-validated ratings improves S/O/D consistency.
No vector DB needed for a small corpus. Ratings still remain DRAFT.

## 7. Integration Points

`llm.draft_scenario(description, model, endpoint, scenarios_dir, use_rag) -> dict`.
Flow: CLI parses `--llm-generate` → `draft_scenario()` → (optional) RAG retrieve →
build prompt → call runtime → validate+annotate → write YAML + `.audit.log` → exit
(does NOT render). Engineer edits the draft, then runs `pfmea.py generate <draft>` as
normal. Lazy import so missing `httpx`/`jsonschema` never affects core users.

## 8. Testing

Mock-endpoint test (no model/GPU/network in CI): patch the HTTP POST to return a minimal
valid scenario; assert `_llm_draft` set and all S/O/D in range. Unit test S/O/D clamping
(15→10, 0→1). Self-check assert at end of `draft_scenario()`. `pytest` completes <2 s.

## 9. Risks & Mitigations

| Risk | Likelihood | Mitigation |
|---|---|---|
| Hallucinated failure modes | High | draft watermark + mandatory review gate + audit log |
| Over-trust in draft S/O/D | High | non-suppressible CLI warning; `DRAFT` comment on every rating; clamp but never silently accept |
| LLM returns prose not JSON | Medium | `format: json` + schema validation; surface raw response |
| Model unavailable/cold | Medium | pre-flight `GET /api/tags`; clear setup error |
| FDA/ISO auditability concern | Medium | audit sidecar is the traceability artifact; output is *input to*, not output of, the formal FMEA |
| Treating draft as final | Low-Med | visible `status: DRAFT` rendered on the workbook until manually set to `REVIEWED` |

**Regulatory note:** a process-efficiency aid, not a validated tool under 21 CFR Part 11.
The drafted YAML equals a hand-typed first draft; the engineer's review and signature
constitute the formal pFMEA record. Document this in the Software Development Plan.
