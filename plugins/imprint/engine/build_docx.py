#!/usr/bin/env python3
"""Kept so existing scripts keep working; the command is now `imprint docx`."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from imprint.build_docx import main  # noqa: E402

raise SystemExit(main())
