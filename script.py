import firebase_admin
from firebase_admin import credentials
from firebase_admin import firestore


from xlrd import open_workbook

cred = credentials.Certificate("path/to/service_account.json")
firebase_admin.initialize_app(cred)

db = firestore.client()


wb = open_workbook("path/to/filename.xlsx")

values = []
for sheet in wb.sheets():
    for row in range(1, sheet.nrows):
        col_names = sheet.row(0)
        col_value = {}
        for name, col in zip(col_names, range(sheet.ncols)):
            value = sheet.cell(row, col).value
            try:
                value = str(value)
            except:
                pass
            col_value[name.value] = value
        db.collection("CollectionName").add(col_value)
        values.append(col_value)


