#!/usr/bin/env bash

set -euo pipefail

export LC_ALL=en_US.utf8
python3 -m pip install briefcase
rm -rf build/
rm -rf dist/
rm -rf logs/
python3 -m briefcase create macOS Xcode
python3 -m briefcase build macOS Xcode
python3 -m briefcase package macOS Xcode -i 0FACCD5E48260DDA00C85D517BCAD13F022A28F1
