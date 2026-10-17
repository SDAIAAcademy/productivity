# AI Productivity: labs

Lab guides for the SDAIA Academy **AI Productivity (رفع الإنتاجية بالذكاء الاصطناعي)** course. Every lab uses AI applications
(chat assistants, NotebookLM, deep research, custom assistants, agents, app builders), not code. Each guide is an Arabic
web page with the goal, the files, timed steps with copyable prompts, check questions, a stretch task ✱ and folded
instructor notes.

All labs share one case study: **مركز الواحة للخدمات المشتركة** (Al-Waha Shared Services Center), a fictional center that
runs HR, finance, procurement, IT and beneficiary services for six fictional government entities. Undergraduates play a
co-op trainee at the center, and some labs have a student variant.

## Quick start

1. Open the [course page](../index.html) and pick a lab, or open `labs/D{day}_{n}_*.html` directly in a browser.
2. Download the files the lab lists from [`files/`](files/). They are synthetic: every name, number and organisation is
   fictional, and national IDs and mobiles only look real.
3. Keep your copy of [`my_productivity_kit.xlsx`](files/my_productivity_kit.xlsx) open all week (task map, productivity log,
   prompt library).

Free tiers are enough for most labs. Where a feature is limited or unavailable in the Kingdom, the guide names a fallback.

## Labs

| Day | Session | Guide | Main files |
|---|---|---|---|
| 1 | My task map and baseline | [`D1_1_task_map.html`](D1_1_task_map.html) | `my_productivity_kit.xlsx`, `inbox_week43.txt` |
| 1 | Test the edges, then sanitise | [`D1_2_edges_and_safety.html`](D1_2_edges_and_safety.html) | `accommodation_request.docx`, `beneficiary_inquiries_q3.xlsx` |
| 1 | From a vague request to a precise prompt (RICE) | [`D1_3_rice.html`](D1_3_rice.html) | `minutes_portal_committee.pdf` |
| 1 | Context document and prompt library | [`D1_4_context_library.html`](D1_4_context_library.html) | `my_productivity_kit.xlsx` |
| 2 | Clear the inbox | [`D2_1_inbox.html`](D2_1_inbox.html) | `inbox_week43.txt`, `report_excerpt_en.docx`, `glossary_en_ar.xlsx` |
| 2 | From transcript to decisions | [`D2_2_transcript_quotes.html`](D2_2_transcript_quotes.html) | `meeting_transcript_sept_review.txt`, `quote_*.pdf`, `procurement_guide.pdf` |
| 2 | Analyse the inquiries, repair the budget | [`D2_3_inquiries_budget.html`](D2_3_inquiries_budget.html) | `beneficiary_inquiries_q3.xlsx`, `training_budget_2026.xlsx`, `service_kpis_2026.xlsx` |
| 2 | Decision deck in 25 minutes | [`D2_4_decision_deck.html`](D2_4_decision_deck.html) | `quarterly_report_q3_2026.pdf`, `visual_identity_guide.pdf` |
| 3 | Research brief and verification log | [`D3_1_research_brief.html`](D3_1_research_brief.html) | none (deep research in two apps) |
| 3 | Ask our policies (notebook) | [`D3_2_ask_our_policies.html`](D3_2_ask_our_policies.html) | `policy_flexible_work.pdf`, `hr_faq.pdf`, `procurement_guide.pdf`, `lecture_05_service_quality.pdf` |
| 3 | Your custom assistant | [`D3_3_custom_assistant.html`](D3_3_custom_assistant.html) | `hr_faq.pdf`, `policy_flexible_work.pdf` |
| 3 | Redesign the monthly report | [`D3_4_workflow_redesign.html`](D3_4_workflow_redesign.html) | `beneficiary_inquiries_q3.xlsx`, `inbox_week43.txt` |
| 4 | Delegate the weekly bulletin to an agent | [`D4_1_bulletin_agent.html`](D4_1_bulletin_agent.html) | `circulars/` |
| 4 | From idea to specification | [`D4_2_idea_to_spec.html`](D4_2_idea_to_spec.html) | none (BRD, SPEC, mock-ups, AGENTS.md) |
| 4 | Build it | [`D4_3_build_it.html`](D4_3_build_it.html) | the documents from 4.2 |
| 4 | Adoption plan and capstone plan | [`D4_4_adoption_plan.html`](D4_4_adoption_plan.html) | `my_productivity_kit.xlsx` |
| 5 | Capstone: canvas, rubric, presentation | [`capstone.html`](capstone.html) | |

## File pack

| File | What it is | Planted for the labs |
|---|---|---|
| `minutes_portal_committee.pdf` | Steering-committee minutes, 4 pages | 3 decisions, an undecided proposal, a task with no date, an escalation above the committee's limit |
| `beneficiary_inquiries_q3.xlsx` | 458 written inquiries with personal data | 6 duplicate rows, inconsistent channel names, no topic column |
| `quarterly_report_q3_2026.pdf` | Quarterly performance report, 12 pages with charts | Summary says satisfaction 84%, the KPI table says 81% |
| `policy_flexible_work.pdf`, `procurement_guide.pdf`, `hr_faq.pdf` | Center policies | Grounding for the notebook and the custom assistant |
| `quote_a_alofuq.pdf`, `quote_b_naseej.pdf`, `quote_c_rawasi.pdf` | Three supplier quotations | Hidden exclusions, hosting outside the Kingdom, escalation clauses, page references |
| `inbox_week43.txt` | Eight emails | An angry beneficiary with personal data, an English invitation, phishing |
| `report_excerpt_en.docx`, `glossary_en_ar.xlsx` | English excerpt and the center's glossary | Translation with fixed terms |
| `meeting_transcript_sept_review.txt` | Automatic transcript, 40 minutes | A 28-minute untranscribed gap, one undecided item, a corrected figure |
| `training_budget_2026.xlsx` | Training budget workbook | `#REF!`, a short `SUM` range, a hard-coded value, a text number, an inverted ratio |
| `service_kpis_2026.xlsx` | Monthly KPIs per entity, January to September | Charts and trends |
| `visual_identity_guide.pdf` | Short visual identity guide | Rules for images and slides |
| `circulars/` | Six weekly circulars and a bulletin template | A duplicate, a cancellation, one restricted circular |
| `accommodation_request.docx` | Workplace-accommodation request | Personal and health data to remove before upload |
| `lecture_05_service_quality.pdf` | University lecture notes | Student variant of the notebook lab |
| `my_productivity_kit.xlsx` | The participant's kit | Task map, productivity log with totals, prompt library |

Answer keys for instructors are in [`instructor/answer_keys.md`](instructor/answer_keys.md).

## Regenerating the files

The pack is generated, so every participant and every answer key match. The Arabic content lives in
[`tools/texts.py`](tools/texts.py), and [`tools/prepare_files.py`](tools/prepare_files.py) renders it:

```bash
uv run labs/tools/prepare_files.py            # XLSX, DOCX, TXT, PDF and the answer keys
uv run labs/tools/prepare_files.py --no-pdf   # without Chrome
```

PDFs are printed by headless Chrome for correct Arabic shaping. Set `CHROME=/path/to/chrome` if it is not found. The
script warns when a page overflows A4.
