import unittest
import getNotebooks
import queries
import plistlib
import os
import sqlite3


class MyTestCase(unittest.TestCase):
    def test_sanity(self):
        self.assertEqual(1+1, 2, "it's working")

    def test_paths_exist(self):
        self.assertTrue(os.path.exists(os.path.expanduser(getNotebooks.ONENOTE_FULL_SEARCH_PATH)),
                        "database folder not found")
        self.assertTrue(os.path.exists(os.path.expanduser(getNotebooks.ONENOTE_USER_INFO_CACHE)),
                        "database folder not found")
        self.assertRaisesRegexp(ValueError, "invalid literal for.*XYZ'$",
                                int, 'XYZ')

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
    def test_get_all(self):
        getNotebooks.init_wf()
        getNotebooks.get_all()

    def test_create_db(self):
        getNotebooks.init_wf()
        getNotebooks.create_db()

    def test_get_children(self):
        getNotebooks.init_wf()
        getNotebooks.get_children('{F71CBF45-C86C-104C-B90B-147B10D2F615}')
        
    def test_top_down(self):
        getNotebooks.init_wf()
        getNotebooks.top_down('{9E1EACD4-A21C-B947-9E83-C599172503C6}{75}')

    def test_search_all_db_entries(self):
        getNotebooks.init_wf()
        getNotebooks.search_all_db_entries()

    def test_search_recent(self):
        getNotebooks.init_wf()
        getNotebooks.get_recent()

    def test_search_modified(self):
        getNotebooks.init_wf()
        getNotebooks.get_modified()

    def test_get_page_name(self):
        getNotebooks.init_wf()
        getNotebooks.get_page_name('{F79D225A-2386-9B48-9D84-68A8AD85A679}')

    def test_get_recent(self):
        getNotebooks.init_wf()
        getNotebooks.get_recent()

    def test_make_url(self):
        getNotebooks.init_wf()
        getNotebooks.make_url(getNotebooks.get_page_name('{6C26AF5F-0E1A-3447-8193-21E36B47A09B}'))

    def test_open_url(self):
        getNotebooks.init_wf()
        item = getNotebooks.NotebookItem(getNotebooks.get_row_by_guid('{8392576E-34AB-2F40-A1F3-46B5ECF3E17A}'))
        url = getNotebooks.make_url(item)
        getNotebooks.open_url(url)

    def test_get_children(self):
        getNotebooks.init_wf()
        item = getNotebooks.NotebookItem(getNotebooks.get_row_by_goid('{9E1EACD4-A21C-B947-9E83-C599172503C6}{20}'))
        children = getNotebooks.get_children(item)

    def test_cache_data(self):
        getNotebooks.init_wf()
        getNotebooks.cache_data()

    def test_get_row(self):
        getNotebooks.init_wf()
        self.assertIsInstance(getNotebooks.get_row_by_guid('{6C26AF5F-0E1A-3447-8193-21E36B47A09B}'),
                              sqlite3.Row,
                              "row is not instance of sqlite.Row")

    class ExpectedFailureTestCase(unittest.TestCase):
        @unittest.expectedFailure
        def test_open_url_fail(self):
            with self.assertRaises(Exception, getNotebooks.open_url(None)) as cm:
                the_exception = cm.exception
                print(the_exception)

        def test_invalid_uid(self):
            self.assertRaises(Exception, getNotebooks.set_user_uid("invalid/path"))


if __name__ == '__main__':
    unittest.main()
