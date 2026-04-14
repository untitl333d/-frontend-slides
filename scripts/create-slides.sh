#!/usr/bin/env bash
# create-slides.sh — Bootstrap a new Reveal.js slide presentation
# Usage: ./scripts/create-slides.sh [output-dir] [theme] [title]
#
# Arguments (all optional):
#   output-dir  Directory to create (default: slides)
#   theme       Reveal.js theme name (default: black)
#   title       Presentation title (default: My Presentation)
#
# Examples:
#   ./scripts/create-slides.sh
#   ./scripts/create-slides.sh slides moon "API Deep Dive"

set -euo pipefail

OUTDIR="${1:-slides}"
THEME="${2:-black}"
TITLE="${3:-My Presentation}"
DATE="$(date +%Y-%m-%d)"

echo "Creating presentation in: $OUTDIR"
echo "Theme: $THEME | Title: $TITLE"

mkdir -p "$OUTDIR"/{css,js,img}

# ── index.html ──────────────────────────────────────────────────────────────
cat > "$OUTDIR/index.html" << HTML
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${TITLE}</title>

  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/reveal.js@5/dist/reveal.css">
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/reveal.js@5/dist/theme/${THEME}.css" id="theme">
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/reveal.js@5/plugin/highlight/monokai.css">
  <link rel="stylesheet" href="css/custom.css">
</head>
<body>
<div class="reveal">
  <div class="slides">

    <!-- ── Title slide ─────────────────────────────────────────────────── -->
    <section>
      <h1>${TITLE}</h1>
      <p><small>${DATE}</small></p>
      <aside class="notes">Welcome and intro.</aside>
    </section>

    <!-- ── Agenda ──────────────────────────────────────────────────────── -->
    <section>
      <h2>Agenda</h2>
      <ul>
        <li class="fragment">Background</li>
        <li class="fragment">Main topic</li>
        <li class="fragment">Key insights</li>
        <li class="fragment">Conclusion</li>
      </ul>
    </section>

    <!-- ── Content section ─────────────────────────────────────────────── -->
    <section>
      <section>
        <h2>Topic Overview</h2>
        <p>High-level introduction to the topic.</p>
      </section>
      <section>
        <h2>Detail A</h2>
        <ul>
          <li>Point one</li>
          <li>Point two</li>
          <li>Point three</li>
        </ul>
      </section>
      <section>
        <h2>Detail B</h2>
        <pre><code class="language-javascript" data-trim>
// Example code block
function example(input) {
  return input.toUpperCase();
}
        </code></pre>
      </section>
    </section>

    <!-- ── Two-column comparison ───────────────────────────────────────── -->
    <section>
      <h2>Comparison</h2>
      <div class="cols">
        <div>
          <h4>Option A</h4>
          <ul>
            <li>Pro: simple</li>
            <li>Con: limited</li>
          </ul>
        </div>
        <div>
          <h4>Option B</h4>
          <ul>
            <li>Pro: flexible</li>
            <li>Con: complex</li>
          </ul>
        </div>
      </div>
    </section>

    <!-- ── Conclusion ──────────────────────────────────────────────────── -->
    <section>
      <h2>Key Takeaways</h2>
      <ul>
        <li class="fragment">Takeaway one</li>
        <li class="fragment">Takeaway two</li>
        <li class="fragment">Takeaway three</li>
      </ul>
    </section>

    <!-- ── Thank you ───────────────────────────────────────────────────── -->
    <section>
      <h2>Thank You</h2>
      <p>Questions?</p>
      <aside class="notes">Open the floor for Q&amp;A.</aside>
    </section>

  </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/reveal.js@5/dist/reveal.js"></script>
<script src="https://cdn.jsdelivr.net/npm/reveal.js@5/plugin/notes/notes.js"></script>
<script src="https://cdn.jsdelivr.net/npm/reveal.js@5/plugin/highlight/highlight.js"></script>
<script src="https://cdn.jsdelivr.net/npm/reveal.js@5/plugin/markdown/markdown.js"></script>
<script>
  Reveal.initialize({
    hash: true,
    slideNumber: true,
    progress: true,
    transition: 'slide',
    plugins: [ RevealNotes, RevealHighlight, RevealMarkdown ]
  });
</script>
<script src="js/custom.js"></script>
</body>
</html>
HTML

# ── css/custom.css ───────────────────────────────────────────────────────────
cat > "$OUTDIR/css/custom.css" << CSS
/* ── Custom presentation styles ─────────────────────────────────────────── */

/* Two-column grid helper */
.reveal .cols {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 2rem;
  align-items: start;
  text-align: left;
}

/* Three-column grid helper */
.reveal .cols-3 {
  display: grid;
  grid-template-columns: 1fr 1fr 1fr;
  gap: 1.5rem;
  align-items: start;
  text-align: left;
}

/* Highlighted callout box */
.reveal .callout {
  border-left: 4px solid currentColor;
  padding: 0.8rem 1.2rem;
  border-radius: 4px;
  margin: 1rem 0;
  text-align: left;
  background: rgba(255,255,255,0.08);
}

/* Full-width code blocks */
.reveal pre {
  width: 100%;
  box-shadow: 0 4px 24px rgba(0,0,0,0.4);
}

/* Smaller text helper */
.reveal .small {
  font-size: 0.7em;
}

/* Centered content helper */
.reveal .center {
  text-align: center;
}

/* Icon row helper */
.reveal .icon-row {
  display: flex;
  justify-content: space-around;
  align-items: center;
  font-size: 3rem;
  margin: 1rem 0;
}
CSS

# ── js/custom.js ─────────────────────────────────────────────────────────────
cat > "$OUTDIR/js/custom.js" << JS
// Custom JavaScript for the presentation
// Runs after Reveal.js initializes

Reveal.on('ready', () => {
  // Log slide count for reference
  console.log('Presentation ready:', Reveal.getTotalSlides(), 'slides');
});
JS

echo ""
echo "✓ Created $OUTDIR/index.html"
echo "✓ Created $OUTDIR/css/custom.css"
echo "✓ Created $OUTDIR/js/custom.js"
echo ""
echo "To preview:"
echo "  python3 -m http.server 8080 --directory $OUTDIR"
echo "  open http://localhost:8080"
