#!/usr/bin/env bash
# Install curated Anthropic skills into ~/.claude/skills/
# Run this after cloning or pulling to wire skills into Claude Code.
set -e

REPO="$(cd "$(dirname "$0")" && pwd)/skills"
TARGET="${HOME}/.claude/skills"

mkdir -p "$TARGET"

skills=(
  docx
  mcp-builder
  pdf
  pptx
  webapp-testing
  xlsx
)

for skill in "${skills[@]}"; do
  ln -sfn "$REPO/$skill" "$TARGET/$skill"
  echo "  linked: $skill"
done

echo "Done — ${#skills[@]} skills linked to $TARGET"
