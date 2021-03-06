import re
import os

ONENOTE_USER_INFO_CACHE = "~/Library/Containers/com.microsoft.onenote.mac/" \
                          "Data/Library/Application Support/Microsoft/UserInfoCache/"
ONENOTE_USER_UID = None

ICON_PAGE = 'icons/page.png'
ICON_SECTION = 'icons/section.png'
ICON_NOTEBOOK = 'icons/notebook.png'
ICON_SECTION_GROUP = 'icons/sectiongroup.png'


class NotebookItem:

    def __init__(self, row):
        self.Type = row[str('Type')]
        self.GOID = row[str('GOID')]
        self.GUID = row[str('GUID')]
        self.GOSID = row[str('GOSID')]
        self.ParentGOID = row[str('ParentGOID')]
        self.GrandparentGOIDs = row[str('GrandparentGOIDs')]
        self.ContentRID = row[str('ContentRID')]
        self.RootRevGenCount = row[str('RootRevGenCount')]
        self.LastModifiedTime = row[str('LastModifiedTime')]
        self.RecentTime = row[str('RecentTime')]
        self.PinTime = row[str('PinTime')]
        self.Color = row[str('Color')]
        self.Title = row[str('Title')]
        self.last_grandparent = self.GrandparentGOIDs
        self.path = None
        self.icon = None
        self.url = None
        self.set_last_grandparent()
        self.set_url()
        self.set_icon()

    def has_parent(self):
        return self.ParentGOID is not None

    def has_grandparent(self):
        return self.GrandparentGOIDs is not None

    def set_last_grandparent(self):
        if self.has_grandparent():
            if len(self.GrandparentGOIDs) > 50:
                grandparents = self.split_grandparents()
                self.last_grandparent = grandparents[-1]

    def split_grandparents(self):
        new_ids = []
        items = self.GrandparentGOIDs.split('}')
        for i in range(len(items) - 1):
            if i % 2 == 0:
                new_ids.append("{0}}}{1}}}".format(items[i], items[i + 1]))
                i += 1

        return new_ids

    def set_path(self, path):
        self.path = path.replace('.one#', '/')

    def set_icon(self):
        if self.Type == 4:
            self.icon = ICON_NOTEBOOK
        elif self.Type == 3:
            self.icon = ICON_SECTION_GROUP
        elif self.Type == 2:
            self.icon = ICON_SECTION
        else:
            self.icon = ICON_PAGE

    def set_url(self):
        if self.Type == 4:
            self.url = 'onenote:https://d.docs.live.net/{0}/Documents/{1}'.format(get_user_uid(), self.Title)
        else:
            self.url = 'onenote:#page-id={0}&end'.format(self.GUID)


def get_user_uid():
    global ONENOTE_USER_UID
    if ONENOTE_USER_UID is None:
        files = os.listdir(os.path.expanduser(ONENOTE_USER_INFO_CACHE))
        for f in files:
            if 'LiveId.db' in f:
                ONENOTE_USER_UID = re.search('(.*)_LiveId\\.db', f).group(1)
    return ONENOTE_USER_UID
