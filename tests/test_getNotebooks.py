import unittest
import getNotebooks
import plistlib
import os
import sqlite3


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

    def CheckSqliteRowIndex(self):
        con = sqlite3.connect(getNotebooks.MERGED_DB)
        con.row_factory = sqlite3.Row
        row = con.execute("select 1 as a, 2 as b").fetchone()
        self.assertTrue(isinstance(row,
                                   sqlite3.Row),
                        "row is not instance of sqlite.Row")

        col1, col2 = row["a"], row["b"]
        self.assertTrue(col1 == 1, "by name: wrong result for column 'a'")
        self.assertTrue(col2 == 2, "by name: wrong result for column 'a'")

        col1, col2 = row["A"], row["B"]
        self.assertTrue(col1 == 1, "by name: wrong result for column 'A'")
        self.assertTrue(col2 == 2, "by name: wrong result for column 'B'")

        col1, col2 = row[0], row[1]
        self.assertTrue(col1 == 1, "by index: wrong result for column 0")
        self.assertTrue(col2 == 2, "by index: wrong result for column 1")

    def test_make_sql(self):
        getNotebooks.make_sql_script()
            # self.debug()
        # self.assertGreater(len(getNotebooks.ALL_DB_PATHS), 0, "no databases found from microsoft")

    # def test_make_url(self):
    #     getNotebooks.init_wf()
        # getNotebooks.make_url()

    def test_create_db(self):
        getNotebooks.create_db()

    def test_get_sec_pages(self):
        getNotebooks.get_section_pages('Algorithm Design')

    def test_search_all_db_entries(self):
        getNotebooks.init_wf()
        getNotebooks.search_all_db_entries()

    def test_get_page_name(self):
        getNotebooks.init_wf()
        getNotebooks.get_page_name('{F79D225A-2386-9B48-9D84-68A8AD85A679}')

    def test_get_recent(self):
        getNotebooks.init_wf()
        getNotebooks.get_recent()

    def test_make_url(self):
        getNotebooks.init_wf()
        getNotebooks.make_url(getNotebooks.get_page_name('{6C26AF5F-0E1A-3447-8193-21E36B47A09B}'))

if __name__ == '__main__':
    unittest.main()
