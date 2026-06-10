# Skill: Natick Town Meeting Video Production Package

Generate a complete video production package (storyboard, speaking script, PowerPoint slides) for a 2–3 minute learning video about Town Meeting debates.

## Trigger

Someone asks to create video materials about Natick Town Meeting topics (gas leaf blowers, artificial turf, etc.) for a 3rd-grade audience.

## Workflow

### Step 1: Gather Sources

Read these files in `~/Curriculum/`:
- `references.md` — compiled web article summaries with position labels (For/Against/Neutral)
- `video_transcripts.md` — video analysis with debate quotes, arguments, and vote outcomes
- `author.md` — Moderator Jeff Alderson bio

Search web or Telvue for any additional missing coverage.

### Step 2: Write Storyboard

Create `storyboard.md` with:
- Production setup: speaker-to-camera format with slides as visual aids
- Slide flow describing when each slide appears
- Notes on natural delivery (words don't need to match slides exactly)

### Step 3: Write Script

Create `script.md` with:
- Conversational speaking guide — not a word-for-word script
- Slide-by-slide flow with key points to cover
- Tips for 3rd-grade pacing and language

### Step 4: Generate Slides

Create/run `generate_slides.py`:
- Uses python-pptx to create clean slides
- Uses existing orientation template for formatting (town seal top-right)
- Slide dimensions: 10" x 5.625"
- Two-column layouts for pro/con debates
- Vote outcomes displayed prominently with numbers

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
| `script.md` | Speaking script (output) |
| `generate_slides.py` | PowerPoint generator (script) |
| `Town_Meeting_Video_Slides.pptx` | PowerPoint output |
| `regenerate.sh` | Regeneration automation |
| `SKILL.md` | This skill definition |

## Third-Grade Language Rules

- Short sentences
- Define every civic term on first use: Town Meeting, Moderator, moratorium
- Concrete examples over abstract concepts
- No jargon without explanation
- Warm, natural tone
