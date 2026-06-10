# Skill: Natick Town Meeting Video Production Package

Generate a complete video production package (storyboard, 3rd-grade script, PowerPoint slides, reusable regeneration script) for a 2-3 minute learning video about Town Meeting debates.

## Trigger

Someone asks to create video materials about Natick Town Meeting topics (gas leaf blowers, artificial turf, etc.) for a 3rd-grade audience.

## Workflow

### Step 1: Gather Sources

Read these files in `~/Curriculum/`:
- `references.md` — compiled web article summaries with position labels (For/Against/Neutral)
- `video_transcripts.md` — video analysis with debate quotes, arguments, and vote outcomes
- `author.md` — Moderator Jeff Alderson bio
- `Copy of Town Government in Massachusetts Reading.txt` — 3rd-grade curriculum context

Search web or Telvue for any additional missing coverage (e.g., if leaf blower debate has no video, note it as a placeholder).

### Step 2: Write Storyboard

Create `storyboard.md` with:
- 10–12 slides for 2–3 minute runtime (~15s per slide)
- Slide number, runtime, visual description, audio description
- Placeholder notes for missing reference material
- 3rd-grade-accessible explanations of: Town Meeting, Moderator role, motion, debate, vote
- Source cross-reference table mapping slides to source materials

### Step 3: Write Script

Create `script.md` with:
- Full voiceover text matching storyboard slide numbers
- Timestamps aligned to storyboard
- Production notes: tone, key terms to emphasize, language level check
- No words above 3rd-grade reading level except key terms (define in context)

### Step 4: Generate Slides

Create/run `generate_slides.py`:
- Uses python-pptx to create minimal-formatting slides
- Slide dimensions: 13.333" x 7.5" (widescreen 16:9)
- Clean design: navy titles, off-white background, teal accents
- Two-column layouts for pro/con debates
- Vote outcomes displayed prominently with numbers
- Placeholder notes for images to be inserted

### Step 5: Regeneration Script

Create/update `regenerate.sh`:
- Check all source files exist
- Run `generate_slides.py`
- Report generated files

## Dependencies

- Python 3 + python-pptx (`pip3 install python-pptx`)

## Files

| File | Purpose |
|------|---------|
| `references.md` | Source references (input) |
| `video_transcripts.md` | Video analysis (input) |
| `author.md` | Moderator bio (input) |
| `storyboard.md` | Video storyboard (output) |
| `script.md` | Voiceover script (output) |
| `generate_slides.py` | PowerPoint generator (script) |
| `Town_Meeting_Video_Slides.pptx` | PowerPoint output |
| `regenerate.sh` | Regeneration automation |
| `SKILL.md` | This skill definition |

## Placeholder Patterns

When source material is thin or missing:
- `[PLACEHOLDER: Insert photo of ___ here]` in slides
- `**Placeholder Notes**` table in storyboard noting what's missing
- Script marks implied content with context clues (no visible placeholders to audience)

## Third-Grade Language Rules

- Short sentences (15 words or fewer where possible)
- Define every civic term on first use: Town Meeting, Moderator, moratorium, motion, debate
- Concrete examples over abstract concepts
- No jargon without explanation
- Warm, clear tone — like reading a picture book aloud
