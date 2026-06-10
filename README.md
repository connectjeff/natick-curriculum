# Natick Town Meeting Video — Curriculum Project

A video production package for a 2–3 minute learning video that teaches 3rd-grade students how Town Meeting works in Massachusetts.

In Massachusetts, towns are run by a Select Board and either a Town Meeting or a Town Council. At Town Meeting, residents vote directly on local laws, projects, and how money is spent. The video brings this concept to life through two real debates from Natick's Fall 2024 and Fall 2025 Town Meetings, showing students how anyone can speak up, how arguments are debated fairly, and how every vote counts.

## Video Summary

**Title:** *How Does Town Meeting Work? Two Big Decisions in Natick*

The video explains:
- What Town Meeting is and how residents vote on big decisions
- The role of the Town Moderator (Jeff Alderson) as a neutral referee
- **Article 32 (Fall 2024):** The debate over phasing out gas-powered leaf blowers — voted 79–27–1 to refer for further study
- **Article 33 (Fall 2025):** The debate over a three-year moratorium on artificial turf fields — voted 74–28–3 to pass
- What students can learn about democracy: anyone can speak up, debate leads to better decisions, and every vote counts

## Video File

<!--
When the final MP4 is added to this repository, replace the placeholder below.
GitHub natively renders .mp4 files referenced in README with a video player.
-->

**Final video:** `Town_Meeting_Video.mp4` *(not yet recorded — add the MP4 file to this directory and update the link below)*

<video src="./Town_Meeting_Video.mp4" controls width="720" poster="">
  Your browser does not support the video tag. Download the file from <a href="./Town_Meeting_Video.mp4">./Town_Meeting_Video.mp4</a>.
</video>

## Artifacts

| File | Description |
|------|-------------|
| [`storyboard.md`](./storyboard.md) | 12-slide storyboard with timing, visual descriptions, audio descriptions, placeholder notes for missing source material, and a source cross-reference table |
| [`script.md`](./script.md) | Full voiceover script with timestamps aligned to the storyboard, production notes, and a 3rd-grade reading-level check |
| [`Town_Meeting_Video_Slides.pptx`](./Town_Meeting_Video_Slides.pptx) | 12 widescreen PowerPoint slides matching the storyboard layout, formatted with the Town Meeting Orientation template (town seal in top-right corner) |
| [`references.md`](./references.md) | 13 compiled sources covering the gas leaf blower and artificial turf debates, each with key-point summaries and position labels (For / Against / Neutral) |
| [`video_transcripts.md`](./video_transcripts.md) | Transcript analysis of 4 Telvue recordings of Fall 2025 Town Meeting sessions, including full debate quotes, arguments from all sides, and vote outcomes for Article 33 |
| [`author.md`](./author.md) | Official bio of Town Moderator Jeff Alderson from natickma.gov, natickveterans.com, and Natick Report articles |
| [`generate_slides.py`](./generate_slides.py) | Python script that regenerates the PowerPoint slides using python-pptx, pulling the town seal image from the orientation template |
| [`regenerate.sh`](./regenerate.sh) | Shell script that checks source files and regenerates all build artifacts |
| [`SKILL.md`](./SKILL.md) | Reusable skill definition documenting the full workflow for generating video production packages from source Markdown files |


## Regeneration

Run `./regenerate.sh` to regenerate the PowerPoint slides from the Python script. Requires Python 3 with `python-pptx` installed:

```bash
pip3 install python-pptx
./regenerate.sh
```

To update the content, edit the source Markdown files (`references.md`, `video_transcripts.md`) and then regenerate the slides.
