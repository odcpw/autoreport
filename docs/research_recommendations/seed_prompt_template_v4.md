# Seed Recommendation Prompt Template (DE) v4

## Purpose
Generate **one recommendation block per EKAS finding** as **four standalone paragraphs** separated by a blank line. Each paragraph stands on its own and corresponds to a maturity step (low → high) without explicit level labels.

Audience: internal engineers/consultants creating client reports.
Tone: Swiss operational realism, concrete, testable, no hype. Use "ss" not "ß". Use umlauts (ä/ö/ü) in the final text.
Never mention FK/Bradley/ADKAR or FK lever codes/names in the output.

## Input (fill in)
- Finding ID: <ID> (internal only; do not print)
- EKAS element (chapter/section): <EKAS element name>
- FK lever (single, internal only): <EB | SA | GE | MN | KM | AB | MO>
- Finding text (DE): <finding>
- Company context (optional): <context>

## Output format (hard rules)
- Exactly 4 paragraphs per finding, separated by a blank line.
- No headings, no numbering, no "Level" labels, no "Step" language, no IDs, no echo of input labels (ID/EKAS/Finding).
- Each paragraph must be a **standalone recommendation** that can be used by itself at that maturity level.
- Do not rely on any other paragraph. No "diese Matrix", "wie bereits", or similar cross‑references.
- Each paragraph must include a clear **Why → Who → How → Proof → Expected result** flow, written as natural prose:
  - **Why**: brief purpose/benefit aligned to maturity level.
  - **Who**: actor (GL / Linienvorgesetzte / SiBe / ASA-Spezialist / Mitarbeitende).
  - **How**: action + artefact + cadence.
  - **Proof**: specific evidence fields.
  - **Expected result**: what becomes visible or stable.
- Style: 6–9 sentences per paragraph, readable and human.
- Avoid repetitive openers (no "Ziel ist …" / "Damit …" in every paragraph). Vary sentence length.
- Rhythm rule: include at least one short sentence (≤8 words) per paragraph.
- Flow rule: include at least one causal connector per paragraph (weil, dadurch, somit, deshalb).
- Closing rule: vary the last sentence; do not use the same closing formula across paragraphs (avoid repeating "Erwartet wird, dass …").
- Avoid ending a sentence with "wird … abgelegt" or "wird … dokumentiert".

## Actor role guardrails (use verbs that fit)
- GL: Freigeben, Priorisieren, Ressourcen, Entscheiden, Budgetieren.
- Linienvorgesetzte: Umsetzen im Tagesgeschäft, Rundgang, Freigabe pro Auftrag, Teambriefing.
- SiBe: Beobachten vor Ort, Sammeln, Stichproben, Nachführung, Rückmeldung.
- ASA-Spezialist: Analysieren, definieren, anpassen, fachliche Prüfung.
- Mitarbeitende: Anwenden, melden, bestätigen, vorschlagen.

## Fixed paragraph intent (non‑negotiable)
1) **Stabilisieren / Mindestwirksamkeit**: minimal intervention that changes behaviour within 3–6 months. No programs, no workshops, no culture talk.
2) **Systematisieren / Ownership**: roles clarified, rhythm defined, light documentation, repeatable process tied to the FK lever (internally).
3) **Alltag + Feedbackschleifen**: embedded in routine management (Führungsbesprechung/Auswertung). Decisions change, not just communication.
4) **Sichern + Lernen**: always-on learning loop (Review/Auswertung/Anpassung), not conditional triggers.

## Escalation rule (must pass)
Pick ONE escalation dimension and make it clearly stronger across paragraphs (without saying it):
- Scope (Einzelfall → Betrieb),
- Systematic depth (pragmatisch → methodisch),
- Learning loop closure (Massnahme → Wirksamkeitsprüfung),
- Ownership (Vorschrift → Eigeninitiative).

## Evidence specificity (hard rule)
Evidence must include at least ONE concrete field or check, not just "mit Datum" or "abgelegt".
Examples: "Protokoll mit Gefährdungstyp, Ort und STOP-Kategorie", "Checkliste mit 3 Pflichtfeldern", "Pendenzenliste mit Termin und Schliessquote pro Monat".

## Instrument budget (hard rule)
- Paragraph 1: max 2 artefacts
- Paragraph 2: max 3 artefacts
- Paragraph 3: max 3 artefacts
- Paragraph 4: max 2 artefacts

## Cadence diversity (hard rule)
Across the 4 paragraphs, use at least two different cadence types.

## STOP rule
If STOP is mentioned, include one concrete S/T/O example that fits the finding.

## Language constraints
- Use concrete verbs: durchführen, dokumentieren, prüfen, kommunizieren, nachführen, freigeben, umsetzen.
- Avoid abstract nouns and jargon: Excellence, Champion, Innovation, Framework, Best Practice, Ownership, Audit, Controlling.
- No rhetorical questions, no slogans, no metaphors.
- Respect STOP hierarchy when measures are mentioned (S‑T‑O‑P).

## Banned terms (examples)
"umfassend", "ganzheitlich", "sicherstellen", "gewährleisten", "Best Practice", "Framework", "Innovation", "Excellence", "Ownership", "Audit", "Controlling", "MO", "MN", "KM", "EB", "SA", "GE", "AB".

## Quality self‑check (must pass)
- Each paragraph contains actor + instrument + cadence + evidence.
- Each paragraph is independently useful and does not reference earlier text or artefacts not defined inside the paragraph.
- Paragraph 4 is not conditional; it is a regular review/learning loop.
- No EKAS drift: actions stay inside the EKAS element and the finding.
- No jargon, no FK codes, no ID lines.
- If it reads like a checklist, rewrite for flow.

## Output
Produce the four‑paragraph recommendation block now.
