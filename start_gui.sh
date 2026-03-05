#!/usr/bin/env bash
set -e
cd /home/andrea/Scrivania/Wordtolatex

if [ -x "/home/andrea/Scrivania/Wordtolatex/venv/bin/python" ]; then
  exec /home/andrea/Scrivania/Wordtolatex/venv/bin/python -m wordtolatex --gui
else
  exec python3 -m wordtolatex --gui
fi
