"""Compatibility wrapper.

The project originally used `check_for_metada_tiff.py` (typo in filename).
This module provides the expected `check_for_metadata_tiff.py` import path
without breaking existing code.
"""

import sys
from pathlib import Path

# Ensure the project root is on sys.path so `backend.*` imports work even when
# running uvicorn from within the `backend/` folder.
_ROOT = Path(__file__).resolve().parents[1]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from backend.check_for_metada_tiff import ProbeResult, google_maps_link, main, make_qr_png, probe_tiff, to_json_dict

__all__ = ["ProbeResult", "probe_tiff", "to_json_dict", "main", "google_maps_link", "make_qr_png"]


if __name__ == "__main__":
    main()
