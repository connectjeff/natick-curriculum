#!/usr/bin/env zsh
# ============================================================
# Regenerate Natick Town Meeting Video Package
# ============================================================
# This script regenerates all video production materials from
# the source Markdown files in ~/Curriculum/.
#
# Prerequisites: python3 with python-pptx installed
#   pip3 install python-pptx
#
# Usage:
#   ./regenerate.sh          # regenerate everything
#   ./regenerate.sh --check  # check source files exist only
# ============================================================

set -e

DIR="${0:A:h}"
SOURCES=(
    "$DIR/references.md"
    "$DIR/video_transcripts.md"
    "$DIR/author.md"
    "Copy of Town Government in Massachusetts Reading.txt"
)

CHECK_ONLY=0
if [[ "$1" == "--check" ]]; then
    CHECK_ONLY=1
fi

echo "=== Natick Video Package Regenerator ==="
echo ""

# ---- Check sources ----
echo "Checking source files..."
for src in $SOURCES; do
    if [[ -f "$src" ]]; then
        echo "  ✓ $src"
    else
        echo "  ✗ MISSING: $src"
        if [[ $CHECK_ONLY -eq 0 ]]; then
            exit 1
        fi
    fi
done

if [[ $CHECK_ONLY -eq 1 ]]; then
    echo ""
    echo "All sources present. Ready to regenerate."
    exit 0
fi

echo ""

# ---- Regenerate storyboard ----
# The storyboard is currently manually authored.
# To regenerate it from source data, run:
#   python3 "$DIR/generate_storyboard.py"
# (Coming soon — see SKILL.md)

# ---- Regenerate script ----
# The script is currently manually authored.
# To regenerate it from storyboard, run:
#   python3 "$DIR/generate_script.py"
# (Coming soon — see SKILL.md)

# ---- Regenerate slides ----
echo "Generating PowerPoint slides..."
python3 "$DIR/generate_slides.py"
echo ""

# ---- Done ----
echo "=== Done ==="
echo ""
echo "Generated files:"
ls -lh "$DIR"/Town_Meeting_Video_Slides.pptx "$DIR"/storyboard.md "$DIR"/script.md
echo ""
echo "To regenerate: run ./regenerate.sh"
