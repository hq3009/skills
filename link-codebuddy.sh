#!/usr/bin/env bash
# Create symbolic links under the skills directories of CodeBuddy / workbuddy
# for every skill (a directory containing SKILL.md) in this repo.
#
# Discovery is recursive: any directory containing SKILL.md is treated as a
# skill and is never descended into. Third-party skills maintained as
# submodules (e.g. under vendor-skills/) are discovered the same way,
# including collection repos that keep several skills under a skills/
# subfolder.
#
# Collection repos may keep helper directories (e.g. "_shared") referenced by
# their skills via "../" relative paths. Underscore-prefixed sibling
# directories of a discovered skill that contain no SKILL.md are linked as
# well, so those references keep working.
#
# Links in the target directories that point into this repo but no longer
# match a discovered skill/helper are removed (stale links). Existing links
# are removed first (the real directory is untouched) and recreated.
#
# Usage:
#   ./link-codebuddy.sh [repo-root]
#   SKILLS_DIRS="/path/a /path/b" ./link-codebuddy.sh [repo-root]
#
# On Windows (Git Bash / MSYS) real symbolic links additionally require
#   export MSYS=winsymlinks:nativestrict
# plus Developer Mode or admin privileges; otherwise ln -s falls back to
# copying the directory tree.

set -euo pipefail

# Repo root, defaults to the script's own directory
REPO_ROOT="${1:-$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)}"
REPO_ROOT="$(cd "$REPO_ROOT" && pwd)"

# Target skills directories; defaults to CodeBuddy and workbuddy user dirs
if [[ -z "${SKILLS_DIRS:-}" ]]; then
    SKILLS_DIRS=("$HOME/.codebuddy/skills" "$HOME/.workbuddy/skills")
else
    read -r -a SKILLS_DIRS <<< "$SKILLS_DIRS"
fi

# --- Discover skill directories: any directory containing SKILL.md ---------
mapfile -t all_candidates < <(
    find "$REPO_ROOT" \
        \( -name .git -o -name node_modules -o -name __pycache__ \
           -o -name .codebuddy -o -name .workbuddy \) -prune \
        -o -type f -name SKILL.md -print |
    while IFS= read -r f; do dirname "$f"; done |
    sort -u
)

# Never keep a candidate nested inside another candidate (a skill directory
# is a leaf unit).
skill_dirs=()
for d in "${all_candidates[@]}"; do
    nested=false
    for p in "${all_candidates[@]}"; do
        if [[ "$p" != "$d" && "$d"/ == "$p"/* ]]; then nested=true; break; fi
    done
    $nested || skill_dirs+=("$d")
done

if [[ ${#skill_dirs[@]} -eq 0 ]]; then
    echo "There's no skill directory containing SKILL.md under $REPO_ROOT" >&2
    exit 0
fi

# --- Discover helper dirs: underscore-prefixed siblings of a skill dir that
# contain no SKILL.md (e.g. the "_shared" folder of a collection repo) ------
declare -A helper_seen
helper_dirs=()
for d in "${skill_dirs[@]}"; do
    parent="$(dirname "$d")"
    for entry in "$parent"/*/; do
        entry="${entry%/}"
        name="$(basename "$entry")"
        [[ "$name" == _* ]] || continue
        [ -f "$entry/SKILL.md" ] && continue
        [[ -n "${helper_seen[$entry]:-}" ]] && continue
        helper_seen["$entry"]=1
        helper_dirs+=("$entry")
    done
done

echo "Discovered ${#skill_dirs[@]} skill(s):"
printf '  %s\n' "${skill_dirs[@]}"
if [[ ${#helper_dirs[@]} -gt 0 ]]; then
    echo "Discovered ${#helper_dirs[@]} helper directorie(s):"
    printf '  %s\n' "${helper_dirs[@]}"
fi

# --- Link creation and stale-link cleanup per target directory -------------
for target_root in "${SKILLS_DIRS[@]}"; do
    mkdir -p "$target_root"

    # Remove stale links that point into this repo but match nothing discovered
    while IFS= read -r -d '' entry; do
        [ -L "$entry" ] || continue
        target="$(readlink "$entry")"
        case "$target" in
            "$REPO_ROOT"/*) ;;
            *) continue ;;
        esac
        name="$(basename "$entry")"
        known=false
        for d in "${skill_dirs[@]}" "${helper_dirs[@]}"; do
            [[ "$(basename "$d")" == "$name" ]] && { known=true; break; }
        done
        if ! $known; then
            rm "$entry"
            echo "Stale link removed: $entry"
        fi
    done < <(find "$target_root" -maxdepth 1 -print0)

    for source in "${skill_dirs[@]}" "${helper_dirs[@]}"; do
        link_path="$target_root/$(basename "$source")"

        if [ -e "$link_path" ] || [ -L "$link_path" ]; then
            rm -rf "$link_path"
        fi

        ln -s "$source" "$link_path"
        echo "Symbolic link created: $link_path -> $source"
    done
done
