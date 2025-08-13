# Ensure the repository root is on sys.path so that `Notebooks` package imports work.
import sys
from pathlib import Path

# This file lives at <repo>/Notebooks/test/conftest.py
# Parents[2] climbs: test -> Notebooks -> <repo>
REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))
