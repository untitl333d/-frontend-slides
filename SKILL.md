---
name: frontend-slides
description: Create interactive web-based slide presentations. Use when the user wants to build a slide deck, presentation, or slideshow using HTML/CSS/JS. Supports Reveal.js with custom theming, code highlighting, speaker notes, and animations. Trigger on requests like "create slides", "make a presentation", "build a slideshow", "turn this into slides".
---

# Frontend Slides Skill

Create beautiful, interactive slide presentations using Reveal.js and web technologies. This skill sets up a complete presentation project with live preview, custom theming, and deployment support.

## Workflow

Make a todo list for all tasks in this workflow and work through them one by one.

### 1. Gather Requirements

Before writing any code, understand what the user needs:

- **Topic**: What is the presentation about?
- **Content**: Do they have existing content (markdown, text, notes) to convert, or do you need to generate slides?
- **Audience**: Who is the audience? (affects tone, depth, terminology)
- **Slide count**: Roughly how many slides?
- **Visual style**: Dark or light theme? Any brand colors?
- **Special features needed**:
  - Code syntax highlighting
  - Speaker notes
  - Slide transitions/animations
  - Charts or diagrams
  - Embedded images/video
- **Output target**: Local files only, GitHub Pages, Netlify, etc.

If the user provided content or a topic, infer what you can from it and proceed without asking unnecessary questions.

### 2. Analyze Existing Content

If the user has existing content to convert (markdown files, README, docs), read those files first:

```bash
# Find relevant content files
find . -name "*.md" -o -name "*.txt" | head -20
```

Extract the key structure — headings become slide titles, subheadings become sections, bullet points become slide content.

### 3. Set Up Project Structure

Create the presentation directory structure:

```bash
mkdir -p slides/{css,js,img,lib}
```

Expected layout:
```
slides/
├── index.html          # Main presentation file
├── css/
│   └── custom.css      # Custom overrides on top of Reveal.js theme
├── js/
│   └── custom.js       # Optional custom JavaScript
├── img/                # Images and assets
└── lib/                # Local copies of libraries (optional)
```

If a `slides/` directory already exists, work within it. If the user wants slides at the project root, place `index.html` there instead.

### 4. Download Reveal.js via CDN (Preferred)

Use CDN links to avoid large dependency installs. The standard Reveal.js CDN setup in `index.html`:

```html
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{{PRESENTATION_TITLE}}</title>

  <!-- Reveal.js core CSS -->
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/reveal.js@5/dist/reveal.css">
  <!-- Theme: choose one of: black, white, league, beige, sky, night, serif, simple, solarized, moon, dracula -->
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/reveal.js@5/dist/theme/{{THEME}}.css" id="theme">
  <!-- Highlight.js for code -->
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/reveal.js@5/plugin/highlight/monokai.css">
  <!-- Custom overrides -->
  <link rel="stylesheet" href="css/custom.css">
</head>
<body>
  <div class="reveal">
    <div class="slides">
      <!-- SLIDES GO HERE -->
    </div>
  </div>

  <script src="https://cdn.jsdelivr.net/npm/reveal.js@5/dist/reveal.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/reveal.js@5/plugin/notes/notes.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/reveal.js@5/plugin/highlight/highlight.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/reveal.js@5/plugin/markdown/markdown.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/reveal.js@5/plugin/math/math.js"></script>
  <script>
    Reveal.initialize({
      hash: true,
      slideNumber: true,
      progress: true,
      transition: '{{TRANSITION}}', // none, fade, slide, convex, concave, zoom
      plugins: [ RevealNotes, RevealHighlight, RevealMarkdown, RevealMath.KaTeX ]
    });
  </script>
  <script src="js/custom.js"></script>
</body>
</html>
```

**Theme selection guide**:
- `black` or `dracula` — dark technical presentations, code-heavy
- `white` or `simple` — clean, minimal, corporate
- `moon` — dark with blue tones, elegant
- `sky` or `beige` — light and friendly
- `night` — dark, high contrast

**Transition guide**:
- `slide` — classic horizontal
- `fade` — subtle, professional
- `convex` or `concave` — dramatic
- `zoom` — attention-grabbing

### 5. Write Slide Content

Each `<section>` is one slide. Nested `<section>` creates vertical slides (sub-sections).

#### Title Slide
```html
<section>
  <h1>{{TITLE}}</h1>
  <h3>{{SUBTITLE}}</h3>
  <p><small>{{AUTHOR}} &bull; {{DATE}}</small></p>
</section>
```

#### Content Slide with Bullets (use `data-auto-animate` for smooth transitions)
```html
<section data-auto-animate>
  <h2>Key Points</h2>
  <ul>
    <li class="fragment">First point revealed on click</li>
    <li class="fragment">Second point</li>
    <li class="fragment">Third point</li>
  </ul>
</section>
```

#### Two-Column Layout
```html
<section>
  <h2>Comparison</h2>
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:2rem">
    <div>
      <h4>Before</h4>
      <ul><li>Old approach</li></ul>
    </div>
    <div>
      <h4>After</h4>
      <ul><li>New approach</li></ul>
    </div>
  </div>
</section>
```

#### Code Slide
```html
<section>
  <h2>Example</h2>
  <pre><code class="language-javascript" data-trim data-noescape>
function greet(name) {
  return `Hello, ${name}!`;
}
  </code></pre>
  <aside class="notes">
    Speaker note: explain the code here.
  </aside>
</section>
```

#### Image Slide
```html
<section data-background-image="img/background.jpg" data-background-opacity="0.3">
  <h2>Visual Slide</h2>
  <p>Text over an image background</p>
</section>
```

#### Markdown Slide (for text-heavy content)
```html
<section data-markdown>
  <textarea data-template>
## Slide Title

- Point one
- Point two
- Point three

Note: Speaker notes go here after "Note:"
  </textarea>
</section>
```

#### Vertical Section (sub-slides grouped under one topic)
```html
<section>
  <!-- Parent slide (horizontal nav) -->
  <section>
    <h2>Topic Overview</h2>
  </section>
  <!-- Sub-slides (vertical nav) -->
  <section>
    <h2>Detail One</h2>
  </section>
  <section>
    <h2>Detail Two</h2>
  </section>
</section>
```

### 6. Create Custom CSS

Write `slides/css/custom.css` to override theme defaults:

```css
/* Custom slide styles */
.reveal {
  font-size: 38px;
}

/* Accent color — replace with brand color */
.reveal h1,
.reveal h2 {
  color: #4a9eff;
}

/* Code blocks */
.reveal pre {
  width: 100%;
  box-shadow: 0 4px 20px rgba(0,0,0,0.3);
}

/* Two-column helper */
.reveal .cols {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 2rem;
  align-items: start;
}

/* Highlight box */
.reveal .callout {
  background: rgba(74,158,255,0.15);
  border-left: 4px solid #4a9eff;
  padding: 1rem 1.5rem;
  border-radius: 4px;
  margin: 1rem 0;
}

/* Fragment fade-up (nicer than default) */
.reveal .fragment.fade-up {
  transform: translateY(20px);
}
```

### 7. Set Up Local Preview

Check if any HTTP server is available:

```bash
# Check what's available
which python3 && python3 --version
which npx && npx --version
which php && php --version
```

Then serve with the best available option:

```bash
# Option A: Python (almost always available)
cd slides && python3 -m http.server 8080

# Option B: Node.js (if available)
npx serve slides -p 8080

# Option C: PHP (if available)
php -S localhost:8080 -t slides/
```

Tell the user to open `http://localhost:8080` in their browser.

**Keyboard shortcuts to tell the user:**
- `→` / `Space` — next slide
- `←` — previous slide
- `↓` / `↑` — navigate vertical slides
- `f` — fullscreen
- `s` — speaker notes view
- `o` — slide overview
- `Esc` — exit overview
- `?` — show all shortcuts

### 8. Optional: Add package.json for npm workflow

If the user wants a proper npm dev workflow (live reload, local Reveal.js):

```json
{
  "name": "slides",
  "version": "1.0.0",
  "scripts": {
    "start": "npx serve slides -p 8080",
    "dev": "npx live-server slides --port=8080",
    "build": "echo 'Static files ready in slides/'"
  },
  "devDependencies": {
    "live-server": "^1.2.2"
  }
}
```

### 9. Optional: GitHub Pages Deployment

If the user wants to deploy to GitHub Pages:

1. The `slides/` directory or root `index.html` can be served directly
2. In the repo settings, set Pages source to the branch and folder

Add a `.github/workflows/pages.yml`:

```yaml
name: Deploy Slides to GitHub Pages
on:
  push:
    branches: [main]
permissions:
  contents: read
  pages: write
  id-token: write
jobs:
  deploy:
    environment:
      name: github-pages
      url: ${{ steps.deployment.outputs.page_url }}
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/configure-pages@v4
      - uses: actions/upload-pages-artifact@v3
        with:
          path: slides/
      - uses: actions/deploy-pages@v4
        id: deployment
```

### 10. Validate the Presentation

Open the presentation in a browser and verify:

```bash
# Check index.html is valid HTML
python3 -c "
from html.parser import HTMLParser
class V(HTMLParser): pass
v = V()
v.feed(open('slides/index.html').read())
print('HTML parsed OK')
"
```

Walk through these checks mentally:
- [ ] Title slide looks correct
- [ ] All sections have proper headings
- [ ] Code blocks are in the right language
- [ ] Fragment animations make sense
- [ ] Speaker notes are present where needed
- [ ] No broken image references

### 11. Commit and Push

```bash
git add slides/ package.json .github/ 2>/dev/null || git add slides/
git commit -m "Add slide presentation: {{TOPIC}}"
git push -u origin HEAD
```

## Common Patterns

### Counting Slides

A typical 20-minute talk needs ~15–20 slides. Rule of thumb: 1 slide per minute.

### Slide Structure Recipe

For a typical technical talk:
1. **Title** — topic, presenter, date
2. **Agenda** — 4–6 bullet points
3. **Problem / Motivation** — 1–2 slides
4. **Background / Context** — 1–2 slides
5. **Solution / Approach** — 3–5 slides (the meat)
6. **Demo / Example** — 1–2 slides
7. **Results / Benchmarks** — 1–2 slides
8. **Conclusion / Takeaways** — 1 slide
9. **Q&A / Thank You** — 1 slide

### Fragment Order

Use `data-fragment-index` to control reveal order:
```html
<p class="fragment" data-fragment-index="2">Appears second</p>
<p class="fragment" data-fragment-index="1">Appears first</p>
```

### Auto-Animate Between Slides

```html
<section data-auto-animate>
  <h2>Step 1</h2>
  <code>initial code</code>
</section>
<section data-auto-animate>
  <h2>Step 2</h2>
  <code>evolved code</code>
</section>
```

Elements with matching `data-id` attributes animate smoothly between slides.

## Wrap Up

When done, provide the user with:

1. **Summary** of slides created (count, sections, key features used)
2. **How to view**: command to start local server + URL
3. **How to navigate**: keyboard shortcuts cheatsheet
4. **How to present**: tips for speaker notes view (`s` key), fullscreen (`f` key)
5. **Next steps**: deployment options, adding images, exporting to PDF

### Exporting to PDF

Reveal.js supports PDF export via browser print:
1. Open the presentation in Chrome/Chromium
2. Add `?print-pdf` to the URL: `http://localhost:8080/?print-pdf`
3. Open browser Print dialog → Save as PDF
4. Set layout to Landscape, disable headers/footers
