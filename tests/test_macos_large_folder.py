"""
Outlook Desktop MCP - macOS large-folder behaviour tests
========================================================
Folders such as Deleted Items or the inbox can hold tens of thousands of
messages. These tests guard the generated AppleScript so a request for a few
items never asks Outlook to build a list of every message in the folder.

Offline: a FakeBridge captures the script, no live Outlook needed.
"""
import asyncio
import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from outlook_desktop_mcp import server_mac  # noqa: E402


class FakeBridge:
    def __init__(self):
        self.script = ""

    async def run(self, script, timeout=None):
        self.script = script
        return ""


class LargeFolderScriptTests(unittest.TestCase):
    def setUp(self):
        self.original_bridge = server_mac.bridge
        self.fake = FakeBridge()
        server_mac.bridge = self.fake

    def tearDown(self):
        server_mac.bridge = self.original_bridge

    def _script(self, coro):
        asyncio.run(coro)
        return self.fake.script

    def _assert_no_full_list(self, script):
        self.assertNotIn("set allMsgs to messages of folderRef", script)
        self.assertNotIn("item i of allMsgs", script)

    def test_latest_items_are_read_by_index_not_from_a_full_list(self):
        script = self._script(server_mac.list_emails("deleted", 3))
        self._assert_no_full_list(script)
        self.assertIn("set msgCount to count of messages of folderRef", script)
        self.assertIn("set m to message i of folderRef", script)
        self.assertIn("set maxCount to 3", script)
        self.assertIn("if msgCount < maxCount then set maxCount to msgCount", script)

    def test_unread_only_scans_in_chunks_without_a_whole_folder_filter(self):
        script = self._script(server_mac.list_emails("deleted", 5, unread_only=True))
        self._assert_no_full_list(script)
        self.assertNotIn("whose", script)
        self.assertIn("is read of (messages startAt thru endAt of folderRef)", script)
        self.assertIn(f"startAt + {server_mac.UNREAD_SCAN_CHUNK} - 1", script)
        # Stops once enough unread messages have been found.
        self.assertIn("repeat while found < maxCount and startAt <= msgCount", script)
        self.assertIn("exit repeat", script)

    def test_unread_only_still_reads_the_same_fields(self):
        script = self._script(server_mac.list_emails("sent", 5, unread_only=True))
        for needle in ("set snd to sender of m", "set msenderName to name of snd", "count of attachments of m", "time received of m"):
            self.assertIn(needle, script)

    def test_count_is_at_least_one(self):
        script = self._script(server_mac.list_emails("sent", 0))
        self.assertIn("set maxCount to 1", script)

    def test_drafts_are_read_by_index_too(self):
        script = self._script(server_mac.list_drafts(limit=5, offset=2))
        self._assert_no_full_list(script)
        self.assertIn("set msgCount to count of messages of folderRef", script)
        self.assertIn("set m to message i of folderRef", script)
        self.assertIn("set startIdx to 3", script)
        self.assertIn("set endIdx to 7", script)


if __name__ == "__main__":
    unittest.main()
