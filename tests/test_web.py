from __future__ import annotations

import unittest

from rrfq_app.web import _enabled_sources


class WebConfigTest(unittest.TestCase):
    def test_manual_csv_payload_enables_manual_source(self):
        sources = _enabled_sources({"enabled_sources": ["lcsc_web"], "manual_csv": "value,unit_price\nABC,1.00"})

        self.assertEqual(sources, {"lcsc_web", "manual_csv"})

    def test_empty_manual_csv_keeps_selected_sources(self):
        sources = _enabled_sources({"enabled_sources": ["lcsc_web"], "manual_csv": "   "})

        self.assertEqual(sources, {"lcsc_web"})


if __name__ == "__main__":
    unittest.main()
