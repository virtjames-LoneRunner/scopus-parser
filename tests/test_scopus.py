import os
import tempfile
import unittest

import pandas

from scopus_parser.config import BISUScopusConfig
from scopus_parser.scopus import BISUScopus


class ScopusTests(unittest.TestCase):
    def setUp(self):
        self.scopus = BISUScopus(BISUScopusConfig())
        self.scopus.data_list_sources = pandas.DataFrame(
            {
                "Sourcerecord ID": [1, 2, 3],
                "Source Title": [
                    "Journal of Farming",
                    "Rural Health Review",
                    "Computer Science Review",
                ],
                "Categories": [
                    "Agriculture (Q1)",
                    "Public Health (Q2)",
                    "Computer Science (Q1)",
                ],
                "Areas": [
                    "Agricultural and Biological Sciences",
                    "Medicine",
                    "Computer Science",
                ],
                "All Science Journal Classification Codes (ASJC)": [
                    "1100",
                    "2700",
                    "1700",
                ],
                "Publisher": ["Farm Press", "Health Press", "Tech Press"],
                "H index": [50, 30, 100],
                "SJR": [10.0, 5.0, 20.0],
                "SJR Best Quartile": ["Q1", "Q2", "Q1"],
            }
        )
        self.scopus.filtered = None

    def test_keyword_any_across_multiple_fields(self):
        self.scopus.filter_by_keywords(
            ["Source Title", "Categories", "Areas"],
            ["farming", "medicine"],
            "any",
        )
        self.assertEqual(set(self.scopus.filtered["Sourcerecord ID"]), {1, 2})

    def test_keyword_all_across_multiple_fields(self):
        self.scopus.filter_by_keywords(
            ["Categories", "Areas"],
            ["agriculture", "biological"],
            "all",
        )
        self.assertEqual(set(self.scopus.filtered["Sourcerecord ID"]), {1})

    def test_sjr_is_converted_to_a_decimal(self):
        csv_text = (
            "Sourceid;SJR;SJR Best Quartile;Categories;Areas;H index\n"
            "1;104,065;Q1;Agriculture (Q1);Agricultural Sciences;50\n"
        )
        with tempfile.TemporaryDirectory() as directory:
            path = os.path.join(directory, "scimagojr.csv")
            with open(path, "w", encoding="utf-8", newline="") as handle:
                handle.write(csv_text)
            self.scopus.add_scimago_rankings(path)
        self.assertAlmostEqual(self.scopus.filtered.loc[0, "SJR"], 104.065)

    def test_asjc_mapping(self):
        self.scopus.data_asjc = {1100: "Agricultural Sciences"}
        self.assertEqual(
            self.scopus.map_multiple_codes("1100"),
            "Agricultural Sciences",
        )

    def test_sort_and_minimum_filters(self):
        self.scopus.filter_by_minimum("H index", 40)
        self.scopus.sort_shortlist("h_index")
        self.assertEqual(
            self.scopus.filtered["Sourcerecord ID"].tolist(),
            [3, 1],
        )

    def test_config_keeps_h_index_and_asjc_separate(self):
        config = BISUScopusConfig()
        self.assertIn("H index", config.columns_to_keep)
        self.assertIn("All Science Journal Classification Codes (ASJC)", config.columns_to_keep)


if __name__ == "__main__":
    unittest.main()
