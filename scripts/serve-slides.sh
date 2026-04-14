#!/usr/bin/env bash
# serve-slides.sh — Start a local HTTP server for slide preview
# Usage: ./scripts/serve-slides.sh [slides-dir] [port]
#
# Arguments (all optional):
#   slides-dir  Directory containing index.html (default: slides)
#   port        Port to serve on (default: 8080)
#
# Examples:
#   ./scripts/serve-slides.sh
#   ./scripts/serve-slides.sh slides 3000

set -euo pipefail

SLIDES_DIR="${1:-slides}"
PORT="${2:-8080}"

if [ ! -d "$SLIDES_DIR" ]; then
  echo "Error: directory '$SLIDES_DIR' not found." >&2
  echo "Run create-slides.sh first, or pass the correct path." >&2
  exit 1
fi

if [ ! -f "$SLIDES_DIR/index.html" ]; then
  echo "Error: no index.html found in '$SLIDES_DIR'." >&2
  exit 1
fi

echo "Serving slides from: $SLIDES_DIR"
echo "URL: http://localhost:$PORT"
echo "Press Ctrl+C to stop."
echo ""
echo "Keyboard shortcuts:"
echo "  → / Space  Next slide"
echo "  ←          Previous slide"
echo "  ↓ / ↑      Navigate vertical slides"
echo "  f          Fullscreen"
echo "  s          Speaker notes"
echo "  o          Slide overview"
echo "  ?          All shortcuts"
echo ""

# Try servers in order of preference
if command -v npx &>/dev/null; then
  echo "Using: npx live-server (with auto-reload)"
  npx --yes live-server "$SLIDES_DIR" --port="$PORT" --no-browser 2>/dev/null || \
  npx --yes serve "$SLIDES_DIR" -l "$PORT"
elif command -v python3 &>/dev/null; then
  echo "Using: python3 http.server"
  python3 -m http.server "$PORT" --directory "$SLIDES_DIR"
elif command -v python &>/dev/null; then
  echo "Using: python SimpleHTTPServer"
  cd "$SLIDES_DIR" && python -m SimpleHTTPServer "$PORT"
elif command -v php &>/dev/null; then
  echo "Using: php built-in server"
  php -S "localhost:$PORT" -t "$SLIDES_DIR"
elif command -v ruby &>/dev/null; then
  echo "Using: ruby WEBrick"
  ruby -run -ehttpd "$SLIDES_DIR" -p "$PORT"
else
  echo "Error: no HTTP server found." >&2
  echo "Install one of: python3, node/npx, php, ruby" >&2
  exit 1
fi
