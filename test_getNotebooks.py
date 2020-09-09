import unittest
import getNotebooks


class MyTestCase(unittest.TestCase):
    def test_sanity(self):
        self.assertEqual(1+1, 2, "it's working")

    # def test_run_trigger(self):
    #     res = getNotebooks.trigger("testing", "com.kfunderburg.oneNoteNav")

    def test_open_url(self):
        getNotebooks.open_url('onenote:https://d.docs.live.net/9478a1a4ec3795b7/Documents/Cooking/Quick Notes.one')

if __name__ == '__main__':
    unittest.main()
