#!/usr/bin/env python3
"""Kept so existing scripts keep working; the command is now `imprint pptx`."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from imprint.build_pptx import main  # noqa: E402

raise SystemExit(main())
