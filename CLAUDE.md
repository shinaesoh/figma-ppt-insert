# CLAUDE.md

This repository is a single Claude Code skill: `figma-ppt-insert`
([SKILL.md](.claude/skills/figma-ppt-insert/SKILL.md)). It inserts Figma-exported
screen images into an existing PowerPoint deck. Work is organized per project
under `projects/{name}/` (`01_template/`, `02_images/`, `03_output/`,
`insert_log.md`); everything under `projects/` stays local (gitignored).

- User guide (Korean, for non-developers): [README.md](README.md)
- Python: type hints, PEP 8, UTF-8 / LF. Geometry in inches, never pixels.
- Text added to slides (none in v1) uses the Pretendard font.
- Out-of-scope requests go to README's "2차 후보" list rather than into code.
