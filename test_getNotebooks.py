import unittest
import getNotebooks
import plistlib

class MyTestCase(unittest.TestCase):
    def test_sanity(self):
        self.assertEqual(1+1, 2, "it's working")

    # def test_run_trigger(self):
    #     res = getNotebooks.trigger("testing", "com.kfunderburg.oneNoteNav")

    # def test_split(self):
    #     s = "Mac > AppleScript > Guides"
    #     self.assertEqual(s.split(' > '), ['Mac', 'AppleScript', 'Guides'])
    #     # check that s.split fails when the separator is not a string
    #     with self.assertRaises(TypeError):
    #         s.split(2)

    def test_get_child(self):
        names = ['Mac', 'AppleScript', 'Guides']
        onenote_pl = plistlib.readPlist(getNotebooks.ONENOTE_PLIST)
        getNotebooks.get_child(onenote_pl, names)

if __name__ == '__main__':
    unittest.main()
