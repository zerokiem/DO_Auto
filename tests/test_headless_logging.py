"""Regression tests for Windows pythonw.exe, where sys.stdout is None."""
from __future__ import annotations

import io
import unittest

from do_auto.log_capture import FileTeeStream
from webapp.run_manager import LogBroadcaster, _TeeStream


class HeadlessLoggingTests(unittest.TestCase):
    def test_web_tee_broadcasts_when_stdout_is_none(self):
        broadcaster = LogBroadcaster()
        stream = _TeeStream(broadcaster, None)

        self.assertEqual(stream.write("dong log\n"), len("dong log\n"))
        stream.flush()

        self.assertEqual(list(broadcaster.history), ["dong log"])

    def test_file_tee_writes_log_when_stdout_is_none(self):
        log_file = io.StringIO()
        stream = FileTeeStream(None, log_file)

        self.assertEqual(stream.write("noi dung\n"), len("noi dung\n"))
        stream.flush()

        self.assertEqual(log_file.getvalue(), "noi dung\n")


if __name__ == "__main__":
    unittest.main()
