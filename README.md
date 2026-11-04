# AI Skills

This repository collects and manages AI Skills: reusable capabilities, task
templates, workflow ideas, and reference materials organized into a consistent
skill structure.

## Directory Structure

```text
skills/
├── README.md
├── link-codebuddy.ps1              # Windows: create skill links (primary)
├── link-codebuddy.sh               # POSIX/Git Bash: create skill links
├── .gitattributes                  # keeps *.sh checked out with LF endings
├── <skill-name>/                   # own / vendored skills (top level)
│   ├── SKILL.md
│   ├── assets/
│   ├── references/
│   └── scripts/
└── vendor-skills/                  # third-party skills maintained as submodules
    ├── paper-deep-reader-skill/    #   single skill
    └── dailypaper-skills/          #   collection repo (several skills + _shared)
        └── skills/
            ├── daily-papers/
            ├── daily-papers-fetch/
            ├── ...
            └── _shared/            # helper dir referenced via ../ by the skills
```

## What Counts as a Skill

Any directory containing a `SKILL.md` is a skill. Discovery is recursive with
two rules:

- a skill directory is a **leaf unit** — it is never scanned for nested skills;
- underscore-prefixed sibling directories of a skill that contain no
  `SKILL.md` (e.g. `_shared`) are treated as **helper directories** and are
  linked too, so that `../` relative references inside collection repos keep
  working.

This means both top-level skills and skills nested inside submodule collection
repos are picked up automatically.

## Installing Skills (Creating Links)

Skills become available to CodeBuddy and workbuddy through links placed in
their skills directories (by default `~/.codebuddy/skills` and
`~/.workbuddy/skills`).

### Windows (primary)

```powershell
.\link-codebuddy.ps1
```

- Discovers all skills recursively and links each one (named after its
  directory) into every target skills directory.
- Removes stale links that point into this repo but no longer match a
  discovered skill (the real directories are never touched).
- Creating symbolic links requires Developer Mode or admin privileges; if a
  symbolic link cannot be created, a directory **junction** is used as a
  fallback (equivalent for local directories).
- Targets can be overridden: `.\link-codebuddy.ps1 -SkillsDirs "C:\path\A","C:\path\B"`

### Linux / macOS / Git Bash

```bash
./link-codebuddy.sh
# override targets:
SKILLS_DIRS="/path/a /path/b" ./link-codebuddy.sh
```

On Windows Git Bash (MSYS), set `export MSYS=winsymlinks:nativestrict` first,
otherwise `ln -s` silently copies the directory tree. Do not run the `.sh`
script on Windows if the `.ps1` is available. Note that `*.sh` files are
forced to LF endings via `.gitattributes` — CRLF endings break bash.

Restart the host session after linking so newly linked skills are scanned.

## Managing Third-Party Skills (Submodules)

Third-party skills live under `vendor-skills/` as git submodules so they can
be synced with their upstream repos at any time.

```bash
# Fresh clone with all submodules
git clone --recurse-submodules <this-repo-url>

# Or initialize after a plain clone
git submodule update --init --recursive

# Add a new third-party skill repo
git submodule add <upstream-url> vendor-skills/<name>

# Sync one submodule to its latest upstream
git submodule update --remote vendor-skills/<name>

# Sync all submodules
git submodule update --remote
```

After adding or updating a submodule, re-run the link script to pick up new
or changed skills.

## Maintenance Principles

- Each skill should correspond to a clear use case.
- The README serves as a high-level overview and does not need to detail every
  individual skill.
- When adding a new skill, keep naming, structure, and folder conventions
  consistent; prefer the directory name to match the `name` field in its
  `SKILL.md` frontmatter (hosts may register skills by either).
- Templates and reference materials should remain as general and reusable as
  possible.
- Never modify files inside `vendor-skills/` submodules directly; changes
  belong upstream and would be lost on the next `submodule update`.
