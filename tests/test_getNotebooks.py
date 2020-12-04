import unittest
import getNotebooks
import plistlib
import os


class MyTestCase(unittest.TestCase):
    def test_sanity(self):
        self.assertEqual(1+1, 2, "it's working")

    def test_db_search(int,issing):
        getNotebooks.db_test()

    def test_paths_exist(self):
        self.assertTrue(os.path.exists(os.path.expanduser(getNotebooks.ONENOTE_FULL_SEARCH_PATH)),
                        "database folder not found")
        self.assertTrue(os.path.exists(os.path.expanduser(getNotebooks.ONENOTE_USER_INFO_CACHE)),
                        "database folder not found")
        self.assertRegexpMatches(getNotebooks.ONENOTE_USER_INFO_CACHE, "_LiveId\\.db")

    def test_invalid_uid(self):
        self.assertRaises(Exception, getNotebooks.set_user_uid("invalid/path"))

    def test_make_sql(self):
        getNotebooks.make_sql_script()
        # self.debug()
        # self.assertGreater(len(getNotebooks.ALL_DB_PATHS), 0, "no databases found from microsoft")

    def test_create_db(self):
        getNotebooks.create_db()


if __name__ == '__main__':
    unittest.main()
