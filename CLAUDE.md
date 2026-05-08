# VBA Practice for ACCT 628 — Project Brief & Claude Context

## What This Is

A supplementary web-based learning tool for ACCT 628 (Accounting with VBA Macros)
at Texas A&M / Mays Business School. Built by CE Gaitz in collaboration with
Professor Joan Sanders. Deployed at vbaguide.netlify.app via Netlify during development.

**This is a SUPPLEMENT to Canvas course materials — not a replacement.**
Every page must make this clear to students.

---

## Repo & Deployment

- GitHub repo: `628-vba-guide` → https://github.com/Charles-Gaitz/vba-guide-v2.git
- Branch strategy: push directly to `main`
- Build command: `npm run build`
- Publish directory: `dist`
- Run locally: `npm run dev`

## Deployment & Hosting

- Current host: Netlify (development and review only)
- Final host: Texas A&M / Mays Business School university servers (static file hosting)
- Build output: hand off entire `dist/` folder to university IT — no server config needed
- URL strategy: explicit `.html` extensions and full `/src/modules/` paths everywhere
- Do NOT use Netlify-specific features (netlify.toml, _redirects, edge functions)
- Canvas links: `https://[domain]/src/modules/loops.html` etc.

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
    sticky-header.js                 ← sticky header + shrink on mobile scroll
    data-table.js                    ← collapsible table + clipboard copy
    exercise-steps.js                ← expandable step logic
  modules/
    foundations.html                 → /src/modules/foundations.html
    programming-concepts.html        → /src/modules/programming-concepts.html
    variables.html                   → /src/modules/variables.html
    loops.html                       → /src/modules/loops.html
    calculations-and-dates.html      → /src/modules/calculations-and-dates.html
    references.html                  → /src/modules/references.html
    filters.html                     → /src/modules/filters.html
    debugging.html                   → /src/modules/debugging.html
    pseudocode.html                  → /src/modules/pseudocode.html
    practice-project.html            → /src/modules/practice-project.html
    objects.html                     → /src/modules/objects.html
```

### Correct href pattern for all internal links:
- Module pages: `/src/modules/[filename].html`
- Home: `/` or `/index.html`
- Section anchors: `/src/modules/[filename].html#section-id`
- Never use `/modules/[filename]` without `/src/` prefix and `.html` extension

### HTML escaping in code blocks:
- Never use bare `<` or `>` inside `<pre><code>` — always escape as `&lt;` `&gt;`
- Never use bare `&` — escape as `&amp;`
- Prism.js renders escaped versions correctly

---

## Module Order (locked)

| # | Module | URL |
|---|---|---|
| 1 | Macro Foundations | `/src/modules/foundations.html` |
| 2 | Adding Programming Concepts | `/src/modules/programming-concepts.html` |
| 3 | Variables | `/src/modules/variables.html` |
| 4 | Loops | `/src/modules/loops.html` |
| 5 | Calculations and Dates | `/src/modules/calculations-and-dates.html` |
| 6 | Relative vs Absolute References | `/src/modules/references.html` |
| 7 | Filters & Shortcut Keys | `/src/modules/filters.html` |
| 8 | F8 Debugging Practice | `/src/modules/debugging.html` |
| 9 | Pseudocode | `/src/modules/pseudocode.html` |

**Practice Project** — full-width feature card on home page, below the 9-card grid. Not a numbered module. URL: `/src/modules/practice-project.html`

**Objects** — file kept (`/src/modules/objects.html`) but not shown in home grid.

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
  --code-keyword:      #e8a0a0;  /* 6.0:1 on --code-bg ✓ */
  --code-comment:      #9a9090;  /* 5.8:1 on --code-bg ✓ */
  --code-string:       #c8e6a0;  /* 6.8:1 on --code-bg ✓ */
  --code-number:       #f0c080;  /* 7.1:1 on --code-bg ✓ */

  /* Pseudocode blocks */
  --pseudo-bg:         #f5f0eb;
  --pseudo-border:     #c8b89a;
  --pseudo-text:       #3a2a1a;  /* 11.4:1 on --pseudo-bg ✓ */

  /* Syntax boxes — NEW */
  --syntax-bg:         #fce8e8;  /* var(--accent-light) */
  --syntax-border:     #500000;  /* var(--accent) */
  --syntax-text:       #3a0000;  /* 10.8:1 on --syntax-bg ✓ */

  /* Surface */
  --surface:           #ffffff;
  --surface-alt:       #fafafa;
  --border:            #e0dbd5;

  /* Coming Soon */
  --coming-soon-bg:    #f5f5f5;
  --coming-soon-text:  #6b6866;

  /* Multiple choice feedback */
  --mc-correct-bg:     #e8f5e9;
  --mc-correct-border: #2e7d32;
  --mc-correct-text:   #1b5e20;
  --mc-incorrect-bg:   #fce8e8;
  --mc-incorrect-border: #500000;
  --mc-incorrect-text: #3a0000;

  /* Layout */
  --header-height:     72px;
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

### Content Width
- Max content width: **960px** centered (updated from 860px)
- This applies to `.site-main`, `.site-header__inner`, `.site-footer__inner`
- On mobile: full width with `var(--space-4)` side padding

### Responsive Breakpoints

```css
--breakpoint-mobile: 640px;
--breakpoint-tablet: 1024px;
```

---

## Sticky Navigation Behavior (LOCKED)

### Site Header
- Fixed to top, always visible: `position: fixed; top: 0; z-index: 100`
- Contains: title (home link) + subtitle + `.anchor-nav` pills
- `--header-height: 72px` — main content has `padding-top: calc(var(--header-height) + var(--space-12))`
- Mobile scroll: adds `.header-scrolled` class → hides subtitle, reduces padding

### Anchor Nav (inside header)
- Pill-style links in maroon header bar — one per section: Concept, Quick Check, Easy Wins, Sample Data, Practice Problem, Challenge
- Default: `rgba(255,255,255,0.12)` bg, `1px solid rgba(255,255,255,0.25)` border
- Hover: `rgba(255,255,255,0.25)` bg, brighter border
- Font: `var(--text-xs)`, white, 500 weight
- Flex wrap on mobile

### Module Nav (`.module-nav`)
- Fixed to bottom of viewport on desktop (above 640px): `position: fixed; bottom: 0; z-index: 90`
- Normal page flow on mobile (below 640px): `position: static`
- Always contains prev/next module links and All Modules center link
- Links wrapped in `.module-nav__inner` for 960px centering
- `.module-nav__home` styles the center "All Modules" link

---

## Component Classes (locked — never rename)

### Layout & Navigation
| Class | Description |
|---|---|
| `.page-intro` | Page title + module number. First child of `<main>`. |
| `.page-title` | H1 inside `.page-intro`. Module name. |
| `.page-subtitle` | "Module N of 8" label. |
| `.anchor-nav` | Nav pills inside fixed header. |
| `.module-nav` | Prev/Next bar at bottom of every module page. |

### Content Boxes
| Class | Description |
|---|---|
| `.box-reminder` | Green. Canvas prerequisite warning. Top of every module. |
| `.box-tip` | Beige. Inline hints. |
| `.course-tip` | 💡 Beige. Peer-voice callout. One per major section. |
| `.syntax-box` | Maroon-tinted. Loop/concept skeleton structure. |
| `.pseudocode-block` | Tan. Pseudocode ONLY. Never for real VBA. |
| `.code-block` | Dark bg. Actual VBA. Prism-tokenized. |
| `.ai-compare` | Stacked vertical. Sanders top (green border), AI below (beige border). |
| `.ai-caution` | Warm sand bg. Warning callout for AI-related cautions. Reusable in all modules. |
| `.ai-caution-label` | Small uppercase heading inside `.ai-caution`. Use `⚠ Cautions: Relying on AI`. |
| `.challenge-framing` | Italic muted italic text placed directly after the Challenge `<h2>`. Reusable in all modules. |

### Interactive
| Class | Description |
|---|---|
| `.mc-question` | Multiple choice block. Lock on click. Immediate feedback. |
| `.mc-option` | A/B/C/D button. `data-correct="true"` on correct answer. |
| `.mc-option.correct` | Green state after reveal. |
| `.mc-option.incorrect` | Red state on wrong selection. |
| `.mc-explanation` | Hidden div. Shown after answer locked. |
| `.quick-check` | Wraps all `.mc-question` blocks in a section. |

### Exercise Cards
| Class | Description |
|---|---|
| `.easy-win` | Tier 1. Simple format or steps. |
| `.exercise-steps` | Expandable step-by-step container. |
| `.exercise-step` | One step. `.open` = expanded. |
| `.exercise-simple` | Simple format: description + hint + solution. |
| `.exercise-hint` | Show/hide hint. |
| `.exercise-solution` | Show/hide solution. |
| `.sample-data-exercise` | Tier 2. Aggie Advisors data. |
| `.exam-challenge` | Tier 3. No hints. Exam-level. |
| `.data-table-section` | Collapsible data table wrapper. |
| `.data-table-toggle` | "Show/Hide Data Table" button. |
| `.copy-data-btn` | Copies TSV to clipboard. |
| `.data-table-wrap` | Hidden by default. Contains table + copy button. |
| `.coming-soon` | Greyed card. No link. |

---

## .syntax-box Component (NEW — LOCKED)

Used in concept sections to show the skeleton structure of a loop or syntax pattern
BEFORE the full code example. Distinct from `.pseudocode-block` (which is for logic)
and `.code-block` (which is for runnable VBA).

```html
<div class="syntax-box">
  <p class="syntax-label">Syntax</p>
  <pre>For [counter] = [start] To [end]
    ' your code here
Next [counter]</pre>
</div>
```

CSS in modules.css:
```css
.syntax-box {
  background: var(--syntax-bg);
  border-left: 4px solid var(--syntax-border);
  border-radius: 4px;
  padding: var(--space-3) var(--space-4);
  margin: var(--space-3) 0 var(--space-4) 0;
  font-family: var(--font-code);
  font-size: var(--text-sm);
  color: var(--syntax-text);
}
.syntax-box pre {
  margin: 0;
  white-space: pre;
  line-height: var(--leading-code);
}
.syntax-label {
  font-size: var(--text-xs);
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.08em;
  color: var(--accent);
  margin: 0 0 var(--space-2) 0;
}
```

---

## Concept Section Content Rules (LOCKED)

1. **Explanation before code** — always. Never drop a code block without preceding prose.
2. **Per loop/concept structure:**
   - 1–2 tight prose paragraphs (no filler — every sentence must add information)
   - `.syntax-box` showing the skeleton structure with placeholders
   - Introduction sentence before the full code example
   - `.code-block` with the full example
3. **Cut filler phrases** — never use:
   - "Think of it as..."
   - "This might sound confusing at first..."
   - "As long as [restatement of what was just said]..."
   - Sentences that restate the previous sentence in different words
4. **Prose tone:** Direct and informative. Students are accounting professionals
   learning a technical skill — treat them as capable adults.
5. **Content width is 960px** — slightly wider paragraphs mean less vertical scroll.
   Keep paragraphs to 3–4 sentences max.

---

## Exercise Format Rules (LOCKED)

**Steps format** — use when: multiple distinct actions in sequence, first time a concept appears
**Simple format** — use when: single observation, concept already introduced above
CONTENT_SPEC.md specifies format per exercise per module.

---

## Locked Anchor IDs (never change)

Every module: `#concept` `#quick-check` `#easy-wins` `#sample-data` `#practice-problem` `#challenge`
- Backward-compat: place `<span id="exam-challenge" aria-hidden="true"></span>` as first child of the `#challenge` section so old deep links still resolve
Practice Project: `#module-1` through `#module-8`, `#data-table`

---

## Module Template Structure (every module — no exceptions)

```
<header>  fixed. title + subtitle + .anchor-nav pills

<main>
  .page-intro          H1 module title + "Module N of 8"
  .box-reminder        Green. Canvas prerequisite.

  #concept
    <h2>
    prose (1–2 tight paragraphs per concept — NO filler)
    .course-tip
    [per-concept: prose → .syntax-box → intro sentence → .code-block]
    .ai-compare        Stacked. Where applicable.

  #quick-check
    .mc-question × 3–5
    .course-tip

  #easy-wins
    .exercise-steps or .exercise-simple (per CONTENT_SPEC.md)

  #sample-data
    .data-table-section   collapsed, copy button

  #practice-problem
    .sample-data-exercise

  #challenge
    <span id="exam-challenge" aria-hidden="true">   ← backward-compat anchor
    .challenge-framing                               ← italic exam framing sentence
    .exam-challenge

.module-nav   ← Prev | All Modules | Next →

<footer>
```

---

## JS Files — module.js import order (locked)

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

## Accessibility (strict university requirement)

- All token colors pre-verified WCAG AA — do not change without re-checking
- Every image: `alt` required
- Every interactive element: visible `:focus` outline using `--accent`
- `aria-expanded` on all toggles
- `aria-disabled="true"` on locked MC options
- Tab order logical on every page
- Lighthouse audit before any page marked complete
- Mobile: all interactive elements minimum 44px touch target

---

## What NOT To Do

- Never use blue tones anywhere
- Never hardcode colors outside tokens.css
- Never deviate from module template structure
- Never put code before explanation prose
- Never use `.pseudocode-block` on real VBA
- Never use `.syntax-box` for runnable code — only for skeleton structures
- Never rename component classes or anchor IDs
- Never invent VBA examples not in CONTENT_SPEC.md
- Never add downloadable files
- Never use Netlify-specific config
- Never write filler prose — cut anything that restates what was just said
- Never include Sub/End Sub wrappers in concept section code examples — show only the lines that demonstrate the concept being taught
- Never include sheet navigation, variable declarations, or setup code in concept examples unless the setup IS the concept being taught
- Comments in code examples must explain the concept, not describe what the line does mechanically (e.g. "moves down one row" on an Offset line is mechanical — cut it; "Always move to next row — inside or outside the IF" explains a concept — keep it)

---

## Build Order

1. ✅ Vite scaffold, tokens.css, base.css, layout.css
2. ✅ Home page
3. ✅ vite.config.js multi-page + all CSS/JS files
4. ✅ Module template + placeholder files
5. ✅ Loops page — reference implementation
6. ✅ Update: content width to 960px, syntax-box styles, loops.html prose trim, sticky nav, header height
7. ✅ Home page: 9-card grid + Practice Project feature card; module sequence corrected
8. ✅ Stub pages created/updated: foundations, programming-concepts, variables, calculations-and-dates, references, filters, debugging, pseudocode — all with correct prev/next nav
9. Variables page content
10. Macro Foundations page content
11. Adding Programming Concepts page content
12. Calculations and Dates page content
13. Relative vs Absolute References page content
14. Filters & Shortcut Keys page content
15. F8 Debugging page content
16. Pseudocode page content
17. Practice Project page content
18. Objects placeholder
19. Polish: Lighthouse, mobile test, cross-links, Canvas URL test