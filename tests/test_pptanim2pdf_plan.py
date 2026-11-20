"""Presentation selection against real local Git histories."""
import importlib.util
import os
from pathlib import Path
import subprocess
import tempfile
import unittest

SCRIPT = Path(__file__).resolve().parents[1] / '.github/scripts/pptanim2pdf_plan.py'
spec = importlib.util.spec_from_file_location('pptanim2pdf_plan', SCRIPT)
planner = importlib.util.module_from_spec(spec)
spec.loader.exec_module(planner)


def git(*args):
    return subprocess.check_output(['git', *args], stderr=subprocess.DEVNULL).decode().strip()


class PresentationPlanTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.previous = Path.cwd()
        os.chdir(self.tmp.name)
        self.addCleanup(self.tmp.cleanup)
        self.addCleanup(os.chdir, self.previous)
        git('init')
        git('config', 'user.name', 'Test')
        git('config', 'user.email', 'test@example.com')
        self.fr = 'Course A/Course_fr.pptx'
        self.en = 'Course A/Course_en.pptx'
        Path('Course A').mkdir()
        for name in (self.fr, self.en, 'exercise.ipynb'):
            Path(name).write_text('source')
        self.commit()

    def commit(self):
        git('add', '-A')
        git('commit', '-m', 'Source change')
        return git('rev-parse', 'HEAD')

    def test_manual_one_source_and_custom_output(self):
        result = planner.plan('workflow_dispatch', input_path=self.fr, output_path='exports/Course_fr.pdf')
        self.assertEqual(result['matrix']['include'], [{'input': self.fr, 'output': 'exports/Course_fr.pdf'}])
        self.assertFalse(result['has_removals'])

    def test_notebook_only_does_not_convert_presentations(self):
        before = git('rev-parse', 'HEAD')
        Path('exercise.ipynb').write_text('changed notebook')
        self.commit()
        self.assertFalse(planner.plan('push', before)['has_conversions'])

    def test_mixed_push_selects_both_changed_presentations(self):
        before = git('rev-parse', 'HEAD')
        for name in (self.fr, self.en, 'exercise.ipynb'):
            Path(name).write_text('changed source')
        self.commit()
        result = planner.plan('push', before)
        self.assertEqual({item['input'] for item in result['matrix']['include']}, {self.fr, self.en})

    def test_rename_removes_old_pdf_and_converts_new_source(self):
        before = git('rev-parse', 'HEAD')
        renamed = 'Course A/New_Course_fr.pptx'
        Path(self.fr).rename(renamed)
        self.commit()
        result = planner.plan('push', before)
        self.assertEqual(result['matrix']['include'], [{'input': renamed, 'output': ''}])
        self.assertEqual(result['removed_matrix']['include'], [{'input': self.fr, 'output': 'Course A/Course_fr.pdf'}])

    def test_tooling_change_rebuilds_all_presentations(self):
        before = git('rev-parse', 'HEAD')
        Path('.github/pptanim2pdf').mkdir(parents=True)
        Path('.github/pptanim2pdf/pptanim2pdf.py').write_text('new converter')
        self.commit()
        result = planner.plan('push', before)
        self.assertEqual({item['input'] for item in result['matrix']['include']}, {self.fr, self.en})

    def test_manual_missing_source_is_rejected(self):
        with self.assertRaisesRegex(ValueError, 'not found'):
            planner.plan('workflow_dispatch', input_path='missing.pptx')


if __name__ == '__main__':
    unittest.main()
