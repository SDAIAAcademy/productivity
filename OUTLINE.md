# رفع الإنتاجية بالذكاء الاصطناعي (AI Productivity): course outline

**SDAIA Academy · AI Fluency Track**: 4 teaching days × 4 sessions + 1 capstone day. Delivered in Arabic.
Each day runs Session 1 → 10' break → Session 2 → 20' break → Session 3 → 10' break → Session 4 (4 × 50 min).
Every session is roughly 20–25 minutes of concepts and live demos followed by a 25-minute lab in real AI applications.

> **From "I tried ChatGPT" to measured, safe, repeatable gains.** Decide what to delegate, describe it precisely,
> check what comes back, and own the result, across writing, documents, data, research, custom assistants,
> agents and apps you build by describing them.

## Who it is for

Government and private-sector employees, and undergraduates, who already use AI apps occasionally and want to use them
well in their actual work or study. No programming. Participants who completed SAMAI (سماي) will recognise the
vocabulary (RICE, the 4D fluency skills, the delegation card); this course goes from awareness to practice.

**Lab environment:** a laptop with a browser and accounts on at least two general assistants (any of ChatGPT, Claude,
Gemini, Microsoft Copilot, HUMAIN Chat), plus NotebookLM and one vibe-coding platform (Etlaq, Lovable, Replit or
Google AI Studio). Free tiers are enough for every lab; each lab names a fallback when a feature needs a paid plan
or is not available in Saudi Arabia. No n8n, Power Automate or code.

## Why this outline (research summary, October 2026)

| Signal | What it means for the course |
|---|---|
| Gains are large on structured, checkable tasks (+14% customer support, +26% software, +50% marketing output; AI Index 2026) but shrink or turn negative outside the "jagged frontier": consultants using AI were 19 points *less* likely to be right on a task outside it (Harvard/BCG). | Day 1 opens with **where AI helps and where it hurts**. Participants map their own tasks against the frontier before learning any tool, and labs include a "trap" task. |
| People misjudge their own speed-up: experienced developers were 19% slower with AI but believed they were 20% faster (METR). Almost all productivity evidence is self-reported. | Every lab writes to a **productivity log** (baseline time vs. AI-assisted time and a quality check). The capstone must show a *measured* gain. No comparable course teaches this. |
| Quality control of AI output is the top AI-era skill (50%) and 86% treat output as a starting point (Microsoft Work Trend Index 2026); 40% of workers received "workslop" last month (BetterUp/Stanford). | Verification is a **practice in every session** and a full session on Day 3 (claim types, lateral reading, source tracing, "would I sign this?"). Labs are judged on verified output, not volume. |
| Saudi workplace AI use is very high (>80% use AI at work, KPMG/Melbourne in AI Index 2026; 69% PwC 2025) but often informal on free tools (Deloitte 2026). SDAIA's Generative AI Guidelines say no classified data in third-party tools; PDPL limits personal data and transfers abroad; NDMO has four classification levels. | **Safe use on Day 1**, before the first file is uploaded: what may go where (share / remove first / never share), masking, pseudonymisation, anonymisation, approved applications. Repeated as a checkpoint in every lab. |
| Prompting guides converge on the same elements (Google TCREI, Microsoft Goal/Context/Source/Expectations, Anthropic Product/Process/Performance); Wharton's research finds prompt "tricks" contingent. SAMAI 2 teaches **RICE** (الدور، التعليمات، السياق، التوقّع). | One scaffold, **RICE**, used all week, plus the habits that matter more than any acronym: state the goal, give the material, define an acceptance criterion, iterate with specific feedback. |
| Individual access alone did not change what work people did (NBER Copilot RCT, 7,137 workers: −2 h/week of email, no task change); organisational factors explain 67% of reported impact (WTI 2026); MIT NANDA reports most pilots without measurable return. | Days 3–4 move from tasks to **workflows**: decompose, decide what to delegate, redesign, measure. The capstone is a redesigned workflow, not a prompt. |
| Information seeking and decision support are among the top uses of assistants (OpenAI/NBER "How People Use ChatGPT"); NotebookLM now produces Arabic audio and video overviews and slide decks. | A **research and knowledge** block: deep research with citation checks, and source-grounded notebooks and projects over the organisation's own documents. |
| Reusable assistants are now standard (custom GPTs, Claude Projects and Skills, Gemini Gems → Skills from 17 Nov 2026, Copilot agents). | Participants build and **test** a custom assistant or skill (typical, messy and edge cases, with and without it). |
| Agents and connectors are mainstream features (ChatGPT agent, Claude Cowork and Claude in Chrome, Copilot Agent Mode in Office, scheduled tasks), but enterprise agent deployment is still in single digits (AI Index 2026) and several features are region-locked. | One **supervised-agent** session: plan → approve → run → review, scoped access, draft-only, "spot the risk". Labs always name a fallback. |
| Non-programmers can now build small tools by describing them (Etlaq, Lovable, Replit, Google AI Studio Build, Claude Artifacts, ChatGPT Sites). | Two **vibe-coding** sessions: idea → business requirements → technical specification → UI mock-ups → build → test → publish, with the prompts SDAIA trainers already use. |
| Novices who delegate their thinking learn less (Anthropic RCT: −17% comprehension; AI Index 2026 "learning penalties"); early-career workers are the heaviest users. | For undergraduates especially: tutor mode, "do it first, then compare", explain-back steps, and a session on AI as a learning partner. |
| No comparable course covers Arabic-specific practice (formal register, dialects, mixed Arabic/English, RTL layouts, Arabic text in images). | Arabic is the default in every lab: formal correspondence, Arabic–English translation with a glossary, checking Arabic output quality and right-to-left layouts. |

Full source list: [README](README.md#sources).

## One case study, all week: **مركز الواحة للخدمات المشتركة**

Al-Waha Shared Services Center (مركز الواحة للخدمات المشتركة) is a fictional government shared-services center: HR,
finance and procurement, communications and beneficiary services for several government entities. Employees play
themselves; undergraduates play a **co-op trainee (متدرّب تعاوني)** joining the center for a semester, and most
labs also offer a study variant. Every file is synthetic, generated by `labs/tools/prepare_files.py`, so labs can be
run on public AI applications without exposing any real person's or organisation's data.

| File pack (`labs/files/`) | Used for |
|---|---|
| Committee minutes: self-service portal launch (PDF) | RICE, extraction, updates for an absent manager |
| Beneficiary inquiries, Q3 (XLSX, ~450 rows with names and mobiles) | Data safety, analysis, charts, assistant test cases |
| Quarterly performance report (PDF, ~16 pages) | Summaries, decision deck, verification |
| Flexible-work policy, procurement guide, HR FAQ (PDF) | Knowledge notebook, custom assistant |
| Three supplier quotations (PDF) | Comparison with page references |
| Inbox export (8 emails) and an English report excerpt with a glossary | Correspondence, translation |
| Meeting transcript, 40 minutes (TXT) | Minutes and action items |
| Training-budget workbook with broken formulas (XLSX) | Spreadsheet work |
| Visual identity guide (PDF) | Images and presentations |
| Weekly circulars folder (DOCX × 6) | Agent and scheduled-task lab |
| Workplace-accommodation request with personal data (DOCX) | Sanitising before upload |

## Through-line: the personal productivity kit (حقيبة الإنتاجية)

Each participant leaves with: a task map and baseline, a work-context document loaded into their assistant, a prompt
library, a productivity log, a source-grounded notebook, a tested custom assistant, a redesigned workflow, an app they
built, and a 30-day adoption plan. Teams combine pieces of it in the capstone.

## Day 1 · Foundations: fluency, limits, safe use and prompting (أساسيات العمل مع الذكاء الاصطناعي)

| Session | Concepts | Lab (25 min) |
|---|---|---|
| **1 · AI at work: where it helps and where it hurts** | What changed in 2025–26; evidence and its caveats (jagged frontier, perception gap, workslop); how an AI application works (application, agent, model, tools; next-token prediction); the landscape incl. Saudi applications; Automation / Augmentation / Agency | **1.1 My task map and baseline**: list 6–8 real tasks, time them, place each on "how closely would I check?", run one task in two assistants side by side, start the productivity log |
| **2 · Fluency (4D), limits and safe use** | التفويض، الوصف، التحقّق، المسؤولية; four machine properties (prediction, knowledge, working memory, steerability) and their failure modes; SDAIA Generative AI Guidelines, PDPL, NDMO levels; share / remove first / never share; masking, pseudonymisation, anonymisation | **1.2 Test the edges, then sanitise**: probe specifics, knowledge cut-off, sycophancy and lost-in-the-middle; sanitise an accommodation request and an inquiries sheet before upload |
| **3 · Prompt engineering with RICE** | What the application guesses when you don't say; RICE; goal vs. format (letter vs. spirit); think of the result first (acceptance criterion); guardrail phrases (اعتمد حصريًا، «غير محدد في المادة») | **1.3 From a vague request to a precise prompt**: committee minutes → update for an absent director; vague vs. RICE; bad-prompt makeover |
| **4 · Conversation, iteration and standing context** | The describe–discern loop; specific feedback; let the AI interview you; examples, steps, "think first"; fresh chats and context limits; custom instructions, memory, styles | **1.4 Context document and prompt library**: the AI interviews you into a work-context document; cold start vs. with context; five reusable templates with placeholders |

## Day 2 · Everyday work: write, read, analyse, present (الإنتاجية في المهام اليومية)

| Session | Concepts | Lab |
|---|---|---|
| **1 · Writing and correspondence in Arabic** | Formal register, tone and audience; replies, letters, memos; rewriting vs. generating; MSA and dialect; translation with a glossary; voice samples; avoiding workslop | **2.1 Clear the inbox**: triage 8 emails, draft replies and one official letter, translate an English excerpt with the center's glossary |
| **2 · Reading, summarising and extracting** | Long documents and attention limits; extraction into tables; transcript → minutes; comparing documents with page references; "quote the evidence" | **2.2 From transcript to decisions**: minutes and action table from a 40-minute transcript; compare three supplier quotations with page references |
| **3 · Data and spreadsheets** | Analysis with code execution; spreadsheet assistants (Copilot, Claude, Gemini in Sheets); formulas and errors explained; let code do the arithmetic; validate on data you already understand | **2.3 Analyse the inquiries**: categorise ~450 inquiries, chart them, reconcile the counts; repair the training-budget workbook |
| **4 · Presentations and visuals** | Report → decision story (one idea per slide, conclusion titles); deck tools (Gamma, Canva, Copilot/Claude in PowerPoint, NotebookLM); RICE for images; Arabic text and RTL; disclosure labels | **2.4 Decision deck in 25 minutes**: quarterly report → 6-slide deck with speaker notes, a cover image and an infographic |

## Day 3 · Research, knowledge and workflows (البحث والمعرفة وإعادة تصميم العمل)

| Session | Concepts | Lab |
|---|---|---|
| **1 · Deep research and verification** | Web search vs. deep research vs. thinking vs. internal knowledge; research briefs; claim types (fabricated reference, unsourced number, quote, overgeneralisation); SIFT and lateral reading; a real 2025 government-report case | **3.1 A verified research brief**: one brief in two deep-research tools; a verification log for five claims |
| **2 · Knowledge you can trust: notebooks and projects** | Source-grounded answers; projects vs. notebooks; how retrieval works (meaning as a place); citations; questions the sources can't answer; audio, video and slide overviews; AI as a tutor | **3.2 Ask our policies**: a NotebookLM notebook over the center's policies; cited answers; trick questions; an Arabic audio overview (students: a lecture notebook with a quiz) |
| **3 · Custom assistants and skills** | Custom GPTs, Claude Projects and Skills, Gemini Gems/Skills, Copilot agents; instructions (product, process, behaviour), knowledge, examples, "never" rules, escalation; testing with and without | **3.3 Build and test an assistant**: an inquiry-reply assistant or a minutes skill, tested on typical, messy and edge cases |
| **4 · Decompose and redesign a workflow** | 1 / 3 / 5-part tasks; the five stations; one application until there is a reason for two; explore → plan → produce → review; AI handles / AI assists, human decides / human only | **3.4 Redesign a monthly report**: chain the steps, time it end to end, compare with the baseline; capstone teams form |

## Day 4 · Agents, vibe coding and adoption (الوكلاء والبرمجة التوليدية والتبنّي)

| Session | Concepts | Lab |
|---|---|---|
| **1 · Agents and scheduled tasks** | The agent loop; chat vs. agent (ChatGPT agent, Claude Cowork and Chrome, Copilot Agent Mode, Manus); connectors and permissions; plan approval; scheduled tasks; working safely (scoped folder, precise verbs, draft-only) | **4.1 Delegate the weekly bulletin**: spot the risk in an agent task, then have an agent turn the circulars folder into a bulletin (draft only) and schedule it |
| **2 · Vibe coding I: from idea to specification** | What vibe coding is; three kinds of apps; the six requirement questions; BRD → technical specification ("interview me first") → UI mock-ups → AGENTS.md; quick prototypes in Claude Artifacts / ChatGPT | **4.2 From idea to spec**: BRD, specification and two UI mock-ups for an internal tool |
| **3 · Vibe coding II: build, test, publish** | Etlaq, Lovable, Replit, Google AI Studio Build; review the plan; one change at a time; versions; six checks before publishing; access levels; no real data | **4.3 Build it**: a working app from the spec, tested by a peer against the acceptance criteria |
| **4 · Responsible adoption + capstone kickoff** | SDAIA's seven AI-ethics principles as workplace decisions; disclosure and the diligence statement; the delegation card; 1 / 7 / 30-day adoption; value indicators and the experiment log | **4.4 Adoption plan and capstone plan**: delegation card, ready prompt, 30-day plan; teams fill the capstone canvas |

## Day 5 · Capstone presentations

Teams of 2–3 redesign one real recurring workflow (their own, or one from the Al-Waha task bank; undergraduates may
pick a study workflow) and present: the baseline, the redesigned workflow and delegation card, the working artefacts
(prompts, custom assistant, agent run or app), a live demo, verification and data checks, the **measured** result of
at least two trial runs, and a 30-day adoption plan. 10 minutes talk + demo, 5 minutes Q&A.

Options: **(1) a tested custom assistant or skill** for a recurring task · **(2) a supervised agent or scheduled
workflow** · **(3) an internal tool built with vibe coding** · **(4) a learning workflow** (undergraduates: a study
assistant with explain-back and quizzes).

| Rubric | Weight |
|---|---|
| Measured impact: baseline vs. trials, time and quality, honest about what didn't work | 25% |
| Workflow design and delegation: what the AI does, what stays human, why | 20% |
| Description: prompts, assistant instructions or app specification | 20% |
| Discernment: acceptance criteria, verification checklist, errors caught | 15% |
| Diligence: data check, approved tools, disclosure, adoption plan | 10% |
| Presentation and live demo | 10% |

## Conventions

- **Vocabulary** follows SAMAI 2 and SDAIA's «معجم البيانات والذكاء الاصطناعي»: الأمر، المخرج، التطبيق، الوكيل،
  النموذج، الأدوات، إطار RICE، التفويض/الوصف/التحقّق/المسؤولية، بطاقة التفويض، معيار القبول، سجلّ التجارب.
  Western digits; English product names as they appear in the applications.
- **Tool names are examples**, not endorsements; features and availability in Saudi Arabia are re-checked the week
  before each cohort. Every lab names a fallback.
- **Attribution:** the 4D framework and several activities are adapted from *AI Fluency: Framework & Foundations*
  (Dakan, Feller and Anthropic, CC BY-NC-SA 4.0).
