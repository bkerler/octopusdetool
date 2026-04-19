#!/usr/bin/env bash

set -euo pipefail

export LC_ALL=en_US.utf8
python3 -m pip install briefcase
rm -rf build/
rm -rf dist/
rm -rf logs/
python3 -m briefcase create macOS Xcode
python3 -m briefcase build macOS Xcode
python3 -m briefcase package macOS Xcode -i E2E71C0E3EB70ECFF463EA07CD45E09D14DD557C
