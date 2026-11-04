import unittest
from unittest.mock import patch
import test_ui_design as design


class LogVisibilityTests(unittest.TestCase):
    setUp = design.DesignTests.setUp

    def test_validation_keeps_user_collapsed_log_for_either_tab(self):
        w=self.window
        w.log_toggle.setChecked(False)
        for context in ('1행 · 실행 전 검증','프로젝트 26012'):
            w._log_validation(dict(status='validated',validation_status='passed',checks=[]),context)
            self.assertFalse(w.log_toggle.isChecked())
            self.assertTrue(w.log.isHidden())
        self.assertIn('검증',w.log.toPlainText())

    def test_visibility_is_restored_when_program_reopens(self):
        w=self.window
        w.log_toggle.setChecked(False)
        other=self.m.AccessManager()
        try:
            self.assertFalse(other.log_toggle.isChecked())
            self.assertTrue(other.log.isHidden())
        finally:
            other.close()
            self.app.processEvents()
