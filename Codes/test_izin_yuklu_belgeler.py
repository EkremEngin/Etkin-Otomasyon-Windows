import unittest
from unittest.mock import Mock

from izin_onaya_dosyali_yildiz import pending_plan_files


class UploadedPdfTests(unittest.TestCase):
    def pending(self, rows, *paths):
        page = Mock()
        page.evaluate.return_value = rows
        return pending_plan_files(page, paths)

    def test_partial_retry_returns_only_missing_files_in_upload_order(self):
        paths = ["/a/İzin.pdf", "/b/İzin (2).pdf", "/a/İzin (3).pdf"]
        self.assertEqual(self.pending([["I\u0307zin (2).pdf"]], *paths), [paths[0], paths[2]])
        self.assertEqual(self.pending([["İzin (3).pdf"], ["İzin.pdf"], ["İzin (2).pdf"]], *paths), [])
        self.assertEqual(self.pending([], *paths), paths)

    def test_unexpected_duplicate_ambiguous_or_unreadable_rows_block(self):
        paths = ["/a/EKİM.pdf", "/a/EKİM (2).pdf"]
        for rows in ([["Başka.pdf"]], [["EKİM.pdf"], ["EKİM.pdf"]], [["EKİM.pdf", "Başka.pdf"]],
                     [["EKİM.pdf", "EKİM (2).pdf"]], [[]], None):
            with self.subTest(rows=rows), self.assertRaises((ValueError, RuntimeError)):
                self.pending(rows, *paths)

    def test_indistinguishable_planned_basenames_block_before_upload(self):
        with self.assertRaises(ValueError):
            self.pending([], "/bir/İzin.pdf", "/iki/I\u0307zin.pdf")
