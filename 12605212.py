import openpyxl

print("PROGRAM STARTED")


workbook = openpyxl.load_workbook("12605212.xlsx", data_only=True)

print("EXCEL FILE OPENED")

sheet = workbook.active

print("Sheet name:", sheet.title)


data = []

for row in sheet.iter_rows(min_row=6, values_only=True):

    if row[0] is None:
            continue

    record = {
         "date": row[0],
         "sleep": row[1],
         "fitness": row[2],
         "study": row[3],
         "coding": row[4],
         "class": row[5],
         "classes_attended": row[6],
         "other": row[7],
         "total_tracked": row[8],
         "free_time": row[9],
         "feeling": row[10],
         "satisfaction": row[11],
         "energy": row[12],
         "notes": row[13]

        }

    data.append(record)

    #display results
    print("total records:", len(data))

    for record in data:
        print(record)
