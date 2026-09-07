#!/usr/bin/env bash
# Create symbolic links under CodeBuddy's skills directory for each skill
# (a directory containing SKILL.md) in the current repo.
#
# Symlinks already existing at the target location are removed first
# (the real source directory is untouched) and recreated.

set -euo pipefail

# Repo root, defaults to the script's own directory
REPO_ROOT="${1:-$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)}"

# CodeBuddy skills directory; defaults to the current user's home path (username generalized)
CODEBUDDY_SKILLS_DIR="${CODEBUDDY_SKILLS_DIR:-$HOME/.codebuddy/skills}"

mkdir -p "$CODEBUDDY_SKILLS_DIR"

found_any=false

for skill_dir in "$REPO_ROOT"/*/; do
    skill_dir="${skill_dir%/}"
    [ -f "$skill_dir/SKILL.md" ] || continue
    found_any=true

    link_path="$CODEBUDDY_SKILLS_DIR/$(basename "$skill_dir")"

    # Handles both regular files/directories and dangling symlinks
    if [ -e "$link_path" ] || [ -L "$link_path" ]; then
        rm -rf "$link_path"
    fi

    ln -s "$skill_dir" "$link_path"
    echo "Symbolic link created: $link_path -> $skill_dir"
done

if [ "$found_any" = false ]; then
    echo "There's no skill directory containing SKILL.md under $REPO_ROOT" >&2
fi
