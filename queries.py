from __future__ import unicode_literals

from getNotebooks import ONENOTE_FULL_SEARCH_PATH, ALL_DB_PATHS


def get_all_items():
    query = "SELECT * FROM Entities;"
    return query


def get_recent_items():
    return "SELECT * FROM Entities ORDER BY RecentTime DESC;"


def get_last_modified():
    return "SELECT * FROM Entities ORDER BY LastModifiedTime DESC;"


def get_parent_row(parent_goid):
    return "SELECT * FROM Entities WHERE GOID = \"{0}\"".format(parent_goid)


def get_row(uid):
    return "SELECT * FROM Entities WHERE GUID = \"{0}\"".format(uid)


def get_children(sec_guid):
    return "SELECT * FROM Entities WHERE ParentGOID = \"{0}\"".format(sec_guid)


def reset_db():
    return "DROP TABLE Entities;"


def create_merged_db():
    for f in os.listdir(os.path.expanduser(ONENOTE_FULL_SEARCH_PATH)):
        if '.db' in f and 'journal' not in f:
            ALL_DB_PATHS.append(ONENOTE_FULL_SEARCH_PATH + f)

    drop_table = "DROP TABLE IF EXISTS Entities;\n"

    create_table = "CREATE TABLE Entities (" \
                   "Type                INTEGER, " \
                   "GOID                NVARCHAR(50) NOT NULL, " \
                   "GUID                NVARCHAR(38) NOT NULL, " \
                   "GOSID               NVARCHAR(50), " \
                   "ParentGOID          NVARCHAR(50), " \
                   "GrandparentGOIDs    TEXT, " \
                   "ContentRID          NVARCHAR(50), " \
                   "RootRevGenCount     INTEGER, " \
                   "LastModifiedTime    INTEGER, " \
                   "RecentTime          INTEGER, " \
                   "PinTime             INTEGER, " \
                   "Color               INTEGER, " \
                   "Title               TEXT, " \
                   "EnterpriseIdentity  TEXT" \
                   ")"

    assert (len(ALL_DB_PATHS) > 0)
    attaches = ""
    inserts = ""
    keys = "Type, GOID, GUID, GOSID, ParentGOID, GrandparentGOIDs, " \
           "ContentRID, RootRevGenCount, LastModifiedTime, RecentTime, " \
           "PinTime, Color, Title, EnterpriseIdentity"

    for i in range(1, len(ALL_DB_PATHS)):
        attaches += "ATTACH DATABASE \"{0}\" as db{1};\n".format(os.path.expanduser(ALL_DB_PATHS[i]), i)
        inserts += "INSERT INTO Entities SELECT {0} FROM db{1}.Entities;\n".format(keys, i)

    sql = drop_table \
          + create_table + "\n\n" \
          + attaches + "\n\n" \
          + inserts

    return sql
