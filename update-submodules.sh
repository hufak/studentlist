#!/usr/bin/env bash
set -euo pipefail

# Initialize public if needed and update it to the remote branch's latest commit.
git submodule update --init --remote --checkout -- public
