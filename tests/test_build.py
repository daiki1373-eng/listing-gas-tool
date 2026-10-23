import importlib.util
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location('build', ROOT / 'scripts/build.py')
build = importlib.util.module_from_spec(spec)
spec.loader.exec_module(build)


class BuildTest(unittest.TestCase):
    def test_checked_in_output_matches_sources(self):
        self.assertEqual(build.bundle(), (ROOT / 'gas/Index.html').read_text(encoding='utf-8'))

    def test_self_contained_html(self):
        html = build.bundle()
        self.assertNotIn('<script src=', html)
        self.assertNotIn('href="style.css"', html)
        self.assertIn('window.createHourlyReport', html)
        self.assertIn('window.ReportHost.ready();', html)
        self.assertIn('<base target="_top">', html)

    def test_templates_are_not_registered_to_a_specific_account(self):
        code = (ROOT / 'gas/Code.gs').read_text(encoding='utf-8')
        self.assertIn("single: ''", code)
        self.assertIn("multi: ''", code)


if __name__ == '__main__':
    unittest.main()
