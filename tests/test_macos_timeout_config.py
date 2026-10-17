"""
Outlook Desktop MCP - macOS AppleScript timeout configuration tests
===================================================================
The per-script and startup timeouts can be raised with the
OUTLOOK_MCP_SCRIPT_TIMEOUT / OUTLOOK_MCP_STARTUP_TIMEOUT environment variables.
Offline: no live Outlook needed.
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from outlook_desktop_mcp.applescript_bridge import _timeout_from_env  # noqa: E402

VAR = "OUTLOOK_MCP_TEST_TIMEOUT"


class TimeoutFromEnvTest(unittest.TestCase):
    def tearDown(self):
        os.environ.pop(VAR, None)

    def test_default_when_unset(self):
        os.environ.pop(VAR, None)
        self.assertEqual(_timeout_from_env(VAR, 30), 30)

    def test_default_when_blank(self):
        os.environ[VAR] = "   "
        self.assertEqual(_timeout_from_env(VAR, 30), 30)

    def test_reads_integer_and_fractional_seconds(self):
        os.environ[VAR] = "120"
        self.assertEqual(_timeout_from_env(VAR, 30), 120.0)
        os.environ[VAR] = "2.5"
        self.assertEqual(_timeout_from_env(VAR, 30), 2.5)

    def test_ignores_bad_values(self):
        for bad in ("abc", "0", "-5", "nan"):
            os.environ[VAR] = bad
            self.assertEqual(_timeout_from_env(VAR, 30), 30, bad)


if __name__ == "__main__":
    unittest.main()
