# External skill collections

## alirezarezvani-claude-skills
- Source: https://github.com/alirezarezvani/claude-skills (commit d3cc4c6, copied 2026-11-27)
- Licence: MIT (see alirezarezvani-claude-skills/LICENSE)
- 501 SKILL.md files, kept in the upstream folder layout.
- Kept outside `.claude/skills/` on purpose: several names clash with the skills already there (internal-comms, brand-guidelines, skill-creator…), and loading 500 skills at once would drown the custom JESA ones. Copy individual skill folders into `.claude/skills/` to activate them.
