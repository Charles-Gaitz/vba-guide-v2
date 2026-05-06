# VBA Practice for ACCT 628 — Project Brief & Claude Context

## What This Is

A supplementary web-based learning tool for ACCT 628 (Accounting with VBA Macros)
at Texas A&M / Mays Business School. Built by CE Gaitz in collaboration with
Professor Joan Sanders. Deployed at vbaguide.netlify.app via Netlify during development.

**This is a SUPPLEMENT to Canvas course materials — not a replacement.**
Every page must make this clear to students.

---

## Repo & Deployment

- GitHub repo: `628-vba-guide`
- Branch strategy: push directly to `main`
- Build command: `npm run build`
- Publish directory: `dist`
- Run locally: `npm run dev`
- Live URL: vbaguide.netlify.app (development only)

## Deployment & Hosting

- Current host: Netlify (development and review only)
- Final host: Texas A&M / Mays Business School university servers (static file hosting)
- Build output: the `dist/` folder is a self-contained static site. Hand off the entire
  `dist/` folder to university IT — no server configuration required.
- URL strategy: use explicit `.html` extensions and full `/src/modules/` paths everywhere.
  Do NOT use Netlify-specific features (netlify.toml, _redirects, edge functions).
  Everything must work on a standard static file server with zero configuration.
- Do NOT add rewrite rules, redirect files, or any host-specific routing config.
- Canvas links use full explicit URLs: `https://[domain]/src/modules/loops.html`

---

## Tech Stack

- Vite (multi-page app — NOT single-page with JS routing)
- Vanilla HTML/CSS/JS (no framework)
- Prism.js for VBA syntax tokenizing — colors overridden entirely by our token CSS
- Each module is a separate .html file with its own direct, linkable URL

---

## URL & File Structure (locked — never deviate)

```
index.html                           → /
src/
  styles/
    tokens.css                       ← ALL CSS variables — single source of truth
    base.css                         ← reset, typography, shared layout
    prism-override.css               ← maps Prism tokens to our color variables
    modules.css                      ← all module component styles
  js/
    quick-check.js                   ← multiple choice logic (shared)
    mobile-toggle.js                 ← ai-compare mobile toggle (shared)
    sticky-header.js                 ← sticky header + shrink on mobile scroll (shared)
    data-table.js                    ← collapsible table + clipboard copy (shared)
    exercise-steps.js                ← expandable step logic (shared)
  modules/
    foundations.html                 → /src/modules/foundations.html
    programming-concepts.html        → /src/modules/programming-concepts.html
    variables.html                   → /src/modules/variables.html
    loops.html                       → /src/modules/loops.html
    references.html                  → /src/modules/references.html
    filters.html                     → /src/modules/filters.html
    debugging.html                   → /src/modules/debugging.html
    practice-project.html            → /src/modules/practice-project.html
    objects.html                     → /src/modules/objects.html
```

### Correct href pattern for all internal links:
- Module pages: `/src/modules/[filename].html`
- Home: `/` or `/index.html`
- Section anchors: `/src/modules/[filename].html#section-id`
- Practice Project cross-links: `/src/modules/practice-project.html#module-N`
- Module cross-links: `/src/modules/[filename].html#exam-challenge`
- Never use `/modules/[filename]` without `/src/` prefix and `.html` extension

### HTML escaping in code blocks:
- Never use bare `<` or `>` inside `<pre><code>` blocks
- Always escape as `&lt;` and `&gt;`
- Prism.js renders escaped versions correctly in the browser

---

## Module Order (locked — used everywhere)

| # | Module | URL |
|---|---|---|
| 1 | Macro Foundations | `/src/modules/foundations.html` |
| 2 | Adding Programming Concepts | `/src/modules/programming-concepts.html` |
| 3 | Variables | `/src/modules/variables.html` |
| 4 | Loops | `/src/modules/loops.html` |
| 5 | Relative vs Absolute References | `/src/modules/references.html` |
| 6 | Filters & Shortcut Keys | `/src/modules/filters.html` |
| 7 | F8 Debugging Practice | `/src/modules/debugging.html` |
| 8 | Practice Project | `/src/modules/practice-project.html` |
| 9 | Objects *(Coming Soon)* | `/src/modules/objects.html` |

---

## Design System (LOCKED — never deviate)

### Colors — TAMU-compliant, all WCAG AA verified

```css
:root {
  /* Primary — TAMU Maroon */
  --accent:            #500000;
  --accent-dark:       #3a0000;
  --accent-light:      #fce8e8;

  /* Text */
  --text:              #1a1a1a;  /* 21:1 on white ✓ */
  --text-muted:        #4a4a4a;  /* 8.9:1 on white ✓ */
  --text-subtle:       #6b6866;  /* 5.5:1 on white ✓ */

  /* Reminder boxes — green */
  --reminder-bg:       #e8f5e9;
  --reminder-border:   #2e7d32;
  --reminder-text:     #1b5e20;  /* 7.2:1 on --reminder-bg ✓ */

  /* Tip boxes — beige */
  --tip-bg:            #fdf6e3;
  --tip-border:        #a0845c;
  --tip-text:          #5c3a1e;  /* 9.3:1 on --tip-bg ✓ */

  /* Code blocks */
  --code-bg:           #1e1e1e;
  --code-text:         #d4d0cb;  /* 10.2:1 on --code-bg ✓ */
  --code-keyword:      #e8a0a0;  /* warm rose — 6.0:1 on --code-bg ✓ */
  --code-comment:      #9a9090;  /* warm gray — 5.8:1 on --code-bg ✓ */
  --code-string:       #c8e6a0;  /* soft green — 6.8:1 on --code-bg ✓ */
  --code-number:       #f0c080;  /* warm amber — 7.1:1 on --code-bg ✓ */

  /* Pseudocode blocks */
  --pseudo-bg:         #f5f0eb;
  --pseudo-border:     #c8b89a;
  --pseudo-text:       #3a2a1a;  /* 11.4:1 on --pseudo-bg ✓ */

  /* Surface */
  --surface:           #ffffff;
  --surface-alt:       #fafafa;
  --border:            #e0dbd5;

  /* Coming Soon cards */
  --coming-soon-bg:    #f5f5f5;
  --coming-soon-text:  #6b6866;  /* 5.5:1 on --coming-soon-bg ✓ */

  /* Multiple choice feedback */
  --mc-correct-bg:     #e8f5e9;
  --mc-correct-border: #2e7d32;
  --mc-correct-text:   #1b5e20;
  --mc-incorrect-bg:   #fce8e8;
  --mc-incorrect-border: #500000;
  --mc-incorrect-text: #3a0000;
}
```

**Hard rules:**
- NO blue tones anywhere
- Never hardcode a color value outside tokens.css
- Never change a token without re-verifying WCAG AA contrast
- Color never conveys meaning alone — pair with icon or text label

### Typography

```css
:root {
  --font-body: 'Inter', system-ui, sans-serif;
  --font-code: 'JetBrains Mono', 'Fira Code', monospace;

  --text-xs:   0.75rem;
  --text-sm:   0.875rem;
  --text-base: 1rem;
  --text-lg:   1.125rem;
  --text-xl:   1.25rem;
  --text-2xl:  1.5rem;
  --text-3xl:  1.875rem;
  --text-4xl:  2.25rem;

  --leading-tight:  1.25;
  --leading-normal: 1.6;
  --leading-code:   1.7;
}
```

### Spacing

```css
:root {
  --space-1:  0.25rem;
  --space-2:  0.5rem;
  --space-3:  0.75rem;
  --space-4:  1rem;
  --space-6:  1.5rem;
  --space-8:  2rem;
  --space-12: 3rem;
  --space-16: 4rem;
}
```

### Responsive Breakpoints

```css
--breakpoint-mobile: 640px;
--breakpoint-tablet: 1024px;
/* Max content width: 860px centered */
```

---

## Sticky Navigation Behavior (LOCKED)

### Site Header
- Fixed to top of viewport on all pages, always visible
- `position: fixed; top: 0; left: 0; right: 0; z-index: 100`
- `<main>` and all page content must have `padding-top` equal to header height
- Desktop: full header with title + subtitle always visible
- Mobile: full header visible at top of page; on scroll down shrinks to title-only
  (subtitle hidden, reduced padding); on scroll back to top expands again
- Shrink behavior handled by `sticky-header.js` adding/removing `.header-scrolled`
  class on `<header>` element

### Anchor Nav
- Sticky just below the fixed header on module pages
- `position: sticky; top: [header-height]px; z-index: 90`
- Always accessible so students can jump to any section without scrolling back up
- On mobile: wraps to two lines if needed, font-size reduces slightly

---

## Component Classes (locked names — never rename)

### Existing components:
| Class | Type | Description |
|---|---|---|
| `.box-reminder` | Green box | Top of every module. Canvas prerequisite. |
| `.box-tip` | Beige box | Inline hints and notes. |
| `.code-block` | Dark bg | Actual VBA. Prism-tokenized. |
| `.pseudocode-block` | Tan bg | Pseudocode ONLY. Never for real VBA. |
| `.ai-compare` | Stacked panels | Sanders on top (green border), AI below (beige border). Always vertical. Tabs hidden on all screen sizes. |
| `.quick-check` | Section | Multiple choice questions. |
| `.easy-win` | Exercise card | Tier 1. |
| `.sample-data-exercise` | Exercise card | Tier 2. Aggie Advisors data. |
| `.exam-challenge` | Exercise card | Tier 3. No hints. |
| `.anchor-nav` | Jump nav | Sticky below header. |
| `.module-nav` | Prev/Next bar | Bottom of every module page. |
| `.course-tip` | Inline callout | Peer-voice. One per major section. |
| `.coming-soon` | Card state | Greyed card. No link. |

### New components (added this update):
| Class | Type | Description |
|---|---|---|
| `.mc-question` | Multiple choice block | Question + options + feedback. |
| `.mc-option` | Button | One answer choice (A/B/C/D). |
| `.mc-option.correct` | State | Green. Applied after reveal. Locked. |
| `.mc-option.incorrect` | State | Red. Applied to wrong selection. Locked. |
| `.mc-explanation` | Hidden div | Shown after answer locked. Explains why. |
| `.data-table-section` | Collapsible container | Wraps the copyable data table. |
| `.data-table-toggle` | Button | "Show Data" / "Hide Data" toggle. |
| `.copy-data-btn` | Button | Copies TSV to clipboard. |
| `.data-table-wrap` | Inner wrapper | Hidden by default, shown on toggle. |
| `.exercise-steps` | Container | Step-by-step guided exercise. |
| `.exercise-step` | Expandable item | One step. Click header to expand. |
| `.exercise-step.open` | State | Step is expanded and visible. |
| `.exercise-simple` | Container | Simple format: description + hint + solution. |
| `.exercise-hint` | Collapsible | Show/hide hint button + content. |
| `.exercise-solution` | Collapsible | Show/hide solution button + content. |

---

## Concept Section Content Rule (LOCKED)

**Explanation always precedes code. Code is illustration, never the lead.**

For Loops specifically: minimum 2 paragraphs of prose per loop type before the code
example appears. The paragraphs must explain:
- What the loop type is and when to use it
- How it works step by step in plain English
- What would happen without it / why it matters

For other modules: minimum 1–2 paragraphs per concept before code appears.
The ratio scales with how abstract the concept is for a non-programmer.

Code blocks must always be preceded by a sentence introducing what the code shows.
Example: "Here is what a For Next loop looks like in practice:" — then the code block.
Never drop a code block without a preceding introduction sentence.

---

## AI Compare Component (stacked vertical — LOCKED)

The `.ai-compare` component is always stacked vertically — Sanders panel on top,
AI panel below. Never side by side. Tabs are hidden on all screen sizes.

- Sanders panel: green left border (`var(--reminder-border)`) — the correct approach
- AI panel: beige left border (`var(--tip-border)`) — shown for contrast only
- Each panel has a small label above the h4: "✅ Sanders Approach" / "⚠️ Typical AI Result"
- `.ai-compare-explanation` sits below both panels in beige tip styling
- Code blocks inside panels use `overflow-x: auto`

---

## Multiple Choice Behavior (LOCKED)

- Each question has A/B/C/D options as `.mc-option` buttons
- Student clicks one option → all options lock immediately (pointer-events: none)
- Correct option gets `.correct` class (green bg, green border, checkmark ✓)
- If student selected wrong: their selection gets `.incorrect` class (red bg)
  AND the correct option gets `.correct` class revealed simultaneously
- `.mc-explanation` div below options is shown after locking
- No retry. No "Check Answers" button. One click, immediate feedback.
- `aria-disabled="true"` added to all options after lock

---

## Data Table Behavior (LOCKED)

- Table is collapsed by default — not visible on page load
- `.data-table-toggle` button: "📋 Show Data Table" → "Hide Data Table" on click
- When expanded, a "Copy to Clipboard" button appears above the table
- Copy button copies data as tab-separated values (TSV) so it pastes into Excel correctly
- Confirmation: button text changes to "✓ Copied!" for 2 seconds then reverts
- Table itself is scrollable horizontally on mobile

---

## Exercise Format Rules (LOCKED)

**Steps format** (`.exercise-steps`) — use when:
- Exercise involves multiple distinct actions in sequence
- Students are new to the concept and need scaffolding
- First time a concept appears in practice

**Simple format** (`.exercise-simple`) — use when:
- Exercise is a single observation or short modification
- Concept has already been introduced in a steps exercise above
- Question is more reflective than procedural

CONTENT_SPEC.md specifies which format each exercise uses per module.

---

## Locked Anchor IDs (never change — Canvas and cross-links depend on these)

### Every module page:
`#concept` `#quick-check` `#easy-wins` `#sample-data` `#exam-challenge`

### Practice Project page:
`#module-1` through `#module-8` and `#data-table`

---

## Module Template Structure (every module — no exceptions)

```
<header class="site-header">   fixed, always visible, shrinks on mobile scroll

<main class="site-main">
  .box-reminder                Canvas prerequisite (green)
  .anchor-nav                  Sticky below header. 5 jump links.

  #concept
    <h2>                       Section heading
    prose paragraphs           2+ paragraphs BEFORE any code (more for abstract concepts)
    .course-tip                "Why this matters" — peer voice
    [per-concept subsections with h3, prose, then code]
    .ai-compare                Always stacked vertical. Where applicable.

  #quick-check
    <h2>
    .mc-question × 3–5        Multiple choice. Lock on click. Immediate feedback.
    .course-tip                "How this shows up on the exam"

  #easy-wins
    <h2>
    .exercise-steps or .exercise-simple   per CONTENT_SPEC.md

  #sample-data
    <h2>
    .data-table-section        Collapsed by default. Copy button. Aggie Advisors data.
    .sample-data-exercise      Exercise using that data.

  #exam-challenge
    <h2>
    .exam-challenge            No hints. No steps. Exam-level. Cross-link to Practice Project.

.module-nav                    Prev | All Modules | Next

<footer>
```

---

## Syntax Highlighting — Prism.js

- Prism.js loaded via CDN on every module page
- Language class: `language-vba` on all VBA `<code>` elements
- Default theme completely overridden by `prism-override.css`
- Never use bare `<` or `>` in code blocks — escape as `&lt;` `&gt;`

---

## JS Files — module.js import order

```js
import './styles/tokens.css';
import './styles/base.css';
import './styles/layout.css';
import './styles/prism-override.css';
import './styles/modules.css';
import './js/sticky-header.js';
import './js/quick-check.js';
import './js/mobile-toggle.js';
import './js/data-table.js';
import './js/exercise-steps.js';
```

---

## Accessibility

- All token colors pre-verified for WCAG AA
- Every image: `alt` required
- Every interactive element: visible `:focus` outline using `--accent`
- `aria-expanded` on all toggles
- `aria-disabled="true"` on locked MC options after answer selected
- Tab order logical on every page
- Lighthouse audit before any page marked complete

---

## Content Rules

- All VBA examples from CONTENT_SPEC.md only — never invent code
- All terminology matches Professor Sanders exactly
- `.pseudocode-block` for pseudocode only — never on actual VBA
- `.course-tip` is peer-voice, never professor-voice
- No downloadable files of any kind
- Concept section: explanation always before code, never after

---

## What NOT To Do

- Never use blue tones anywhere
- Never hardcode colors outside tokens.css
- Never deviate from module template structure
- Never put code before explanation prose in concept sections
- Never use `.pseudocode-block` on real VBA
- Never rename component classes or anchor IDs
- Never invent VBA examples not in CONTENT_SPEC.md
- Never add downloadable files
- Never use Netlify-specific config

---

## Build Order

1. ✅ Vite scaffold, tokens.css, base.css, layout.css, header/footer shell
2. ✅ Home page
3. ✅ vite.config.js multi-page + prism-override.css + modules.css
4. ✅ module-template.html + placeholder module files
5. ✅ quick-check.js (old version) + mobile-toggle.js
6. → Update: sticky-header.js, data-table.js, exercise-steps.js (new)
7. → Update: quick-check.js to multiple choice logic
8. → Update: modules.css with new component styles
9. → Rebuild loops.html with all new components
10. Variables page
11. Macro Foundations page
12. Adding Programming Concepts page
13. F8 Debugging page
14. Relative vs Absolute References page
15. Filters & Shortcut Keys page
16. Practice Project page
17. Objects placeholder page
18. Polish: Lighthouse audit, mobile test, all cross-links verified