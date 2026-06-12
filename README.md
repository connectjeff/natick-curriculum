# Natick Town Meeting Video — Curriculum Project

A video production package for a 2–3 minute learning video that teaches 3rd-grade students how Town Meeting works in Massachusetts.

In Massachusetts, towns are run by a Select Board and either a Town Meeting or a Town Council. At Town Meeting, elected Town Meeting members vote on local laws, projects, and how money is spent — and any resident can come and speak. The video brings this concept to life through two real debates from Natick's Fall 2024 and Fall 2025 Town Meetings, showing students how anyone can speak up, how arguments are debated fairly, and how every vote counts.

## Production Approach

The video features Town Moderator Jeff Alderson speaking directly to camera with a Massachusetts flag visible behind him. Slides appear on screen as visual references as he walks through each topic. The tone is natural and conversational — not a produced voiceover.

## Video Summary

**Title:** *How Does Town Meeting Work? Two Big Decisions in Natick*

The video explains:
- What Town Meeting is and how elected members vote on big decisions
- The role of the Town Moderator (Jeff Alderson) as a neutral referee
- **Article 32 (Fall 2024):** The debate over phasing out gas-powered leaf blowers — voted 79–27–1 to refer for further study
- **Article 33 (Fall 2025):** The debate over a three-year moratorium on artificial turf fields — voted 74–28–3 to pass
- What students can learn about democracy: anyone can speak up, debate leads to better decisions, and every vote counts

## Video Recording

**Final video with slides:** [`Town Meeting Recording with Slides.mp4`](./Town%20Meeting%20Recording%20with%20Slides.mp4)

<video src="./Town%20Meeting%20Recording%20with%20Slides.mp4" controls width="720">
  <track src="./video_transcript.vtt" kind="captions" srclang="en" label="English">
  Your browser does not support the video tag. Download the file from <a href="./Town%20Meeting%20Recording%20with%20Slides.mp4">Town Meeting Recording with Slides.mp4</a>.
</video>

**Transcript:** [`Readable transcript`](./video_transcript.md) | [`VTT captions`](./video_transcript.vtt)

Additional recording exports:
- [`Town Meeting Recording only Video.mp4`](./Town%20Meeting%20Recording%20only%20Video.mp4)
- [`Town Meeting Recording only Slides.mp4`](./Town%20Meeting%20Recording%20only%20Slides.mp4)

## Artifacts

| File | Description |
|------|-------------|
| [`storyboard.md`](./storyboard.md) | Storyboard describing the speaker-to-camera format, slide flow, and production setup |
| [`script.md`](./script.md) | Conversational speaking guide with key points to cover per slide |
| [`Town_Meeting_Video_Slides.pptx`](./Town_Meeting_Video_Slides.pptx) | 11 widescreen PowerPoint slides formatted with the Town Meeting Orientation template (town seal in top-right corner) |
| [`Town Meeting Recording with Slides.mp4`](./Town%20Meeting%20Recording%20with%20Slides.mp4) | Final recorded video with speaker video and slides composited together |
| [`Town Meeting Recording only Video.mp4`](./Town%20Meeting%20Recording%20only%20Video.mp4) | Speaker-only video export |
| [`Town Meeting Recording only Slides.mp4`](./Town%20Meeting%20Recording%20only%20Slides.mp4) | Slides-only video export |
| [`video_transcript.md`](./video_transcript.md) | Readable transcript for the final recorded video |
| [`video_transcript.vtt`](./video_transcript.vtt) | WebVTT caption file extracted from the final recorded video |
| [`Camtasia Project from 12JUN2026.zip`](./Camtasia%20Project%20from%2012JUN2026.zip) | Camtasia project archive used to produce the final recording exports |
| [`references.md`](./references.md) | 13 compiled sources covering the gas leaf blower and artificial turf debates, each with key-point summaries and position labels (For / Against / Neutral) |
| [`video_transcripts.md`](./video_transcripts.md) | Transcript analysis of Fall 2024 and Fall 2025 Town Meeting sessions — Article 32 gas leaf blower debate (79-27-1 referral) and Article 33 artificial turf moratorium (74-28-3 passed), with full debate quotes from all sides and vote outcomes |
| [`auto-generated-r.vtt`](./auto-generated-r.vtt) | Raw auto-generated VTT caption file for the Fall 2024 ATM Session 3 recording, containing the full Article 32 leaf blower debate |
| [`author.md`](./author.md) | Official bio of Town Moderator Jeff Alderson from natickma.gov, natickveterans.com, and Natick Report articles |
| [`generate_slides.py`](./generate_slides.py) | Python script that regenerates the PowerPoint slides using python-pptx, pulling the town seal image from the orientation template |
| [`regenerate.sh`](./regenerate.sh) | Shell script that checks source files and regenerates all build artifacts |
| [`SKILL.md`](./SKILL.md) | Reusable skill definition documenting the full workflow |


## Regeneration

Run `./regenerate.sh` to regenerate the PowerPoint slides from the Python script. Requires Python 3 with `python-pptx` installed:

```bash
pip3 install python-pptx
./regenerate.sh
```

To update the content, edit the source Markdown files (`references.md`, `video_transcripts.md`) and then regenerate the slides.
