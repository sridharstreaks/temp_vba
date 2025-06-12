#!/usr/bin/env bash
set -euo pipefail
shopt -s nullglob

BASE="demo"

for proj in "$BASE"/*/; do
  # strip trailing slash, then get just the project name
  proj="${proj%/}"
  proj_name="${proj##*/}"

  # build an array of its immediate sub‐directories (with trailing slash)
  subdirs=( "$proj"/*/ )
  # if there are no subdirs, skip
  (( ${#subdirs[@]} )) || continue

  # sort them lexically, preserving their full paths
  IFS=$'\n' sorted=( $(printf '%s\n' "${subdirs[@]}" | sort) )
  unset IFS

  # pick the last one
  last_sub="${sorted[-1]}"
  last_name="${last_sub%/}"
  last_name="${last_name##*/}"

  printf 'Project: %q → Last folder: %q\n' "$proj_name" "$last_name"
done
