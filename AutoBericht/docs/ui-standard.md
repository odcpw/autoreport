# AutoBericht UI Standard (Draft)

Purpose: keep the current visual language, while making components consistent across AutoBericht and PhotoSorter.

Implementation: shared styles live in `AutoBericht/mini/shared/ui.css` and are loaded by both apps. Page-specific CSS should only handle layout or unique components.

## Foundations
- **Typeface:** "Avenir", "Segoe UI", system-ui, sans-serif.
- **Base background:** `#f6f3ef`.
- **Primary text:** `#1a232f` / `#1f2933`.
- **Accent (primary button):** `#f6d365`.
- **Surface (panels/cards/modals):** `#fdfaf5`.
- **Borders:** `#e0d6c9` (light), `#efe5dc` (soft panel).
- **Status warning:** background `#fff3cd`, text `#5f4b00`.

## Buttons
- **Primary:** filled, rounded pill (`border-radius: 999px`), `#f6d365` background, dark text.
- **Ghost:** light neutral background (`#ffffff` or `#e7dccf`) with dark text.
- **Disabled:** 60% opacity, no hover.
- **Action rows:** right-aligned with an 8px gap.

## Panels & Cards
- **Panel radius:** 12px.
- **Modal radius:** 16px.
- **Panel borders:** `1px solid #efe5dc`.
- **Shadow:** subtle (`0 6px 16px rgba(0,0,0,0.04)` for panels; stronger for modals).

## Modals & Overlays
Structure:
- Backdrop: `rgba(23, 32, 42, 0.45)`.
- Panel: background `#fdfaf5`, border `1px solid #e0d6c9`, radius 16px, shadow `0 16px 40px rgba(0,0,0,0.2)`.
- Header: title on left, close button on right, optional divider line.
- Body: 12px vertical rhythm; avoid full-width button stacks unless a single CTA.
- Actions: right-aligned; primary action first, secondary ghost.

Recommended class names:
- `.modal`, `.modal__backdrop`, `.modal__content`, `.modal__header`, `.modal__body`, `.modal__actions`

## Forms
- Labels in small uppercase with letter spacing.
- Inputs: 10px radius, `1px solid #d7d0c7`, light background.

## Consistency Targets
- **Modals:** use the same backdrop, radius, border, and shadow across the app.
- **Action rows:** consistent spacing and alignment.
- **Copy tone:** short, direct, task-oriented (1–2 sentences max in modals).
