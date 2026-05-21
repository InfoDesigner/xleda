"""Check that basic features work.

Catch cases where e.g. files are missing so the import doesn't work. It is
recommended to check that e.g. assets are included."""

import sys
from xleda import FieldAnalysis  # noqa: F401



if 'xleda' in sys.modules:
    print("Smoke test succeeded")
else:
    raise RuntimeError("xleda import failed")
