import pandas as pd


# Global Variables
name_map = dict()
name_idx = dict()
data = dict()
all_names = []

def read_data(fileName:str):
    attend = pd.read_excel(fileName, sheet_name="Attendance")

    names = list()

    for name in attend['Attendee']:
        if name not in name_map.keys():
            continue
        real_name = name_map[name]
        names.append(real_name)

    # Fill Found
    for i in range(len(name_idx)):
        if name_idx[i] in names:
            data[i][fileName] = 1
    # Fill Not Found
    for i in range(len(all_names)):
        current_name = all_names[i]
        if current_name in names:
            continue
        data[i][fileName] = 0
    return 


def main():
    
    # Initial State
    all_students = pd.read_excel('students.xlsx', sheet_name="Students")

    for i in range(len(all_students['Name'])):
        current_user = all_students['Username'][i]
        current_name = all_students['Name'][i]
        name_map[current_user] = current_name
        all_names.append(current_name)

    for i in range(len(all_names)):
        data[i] = dict()
        data[i]["Name"] = all_names[i]
        name_idx[i] = all_names[i]

    # Solving Session
    read_data("20_3_Afternoon.xlsx")
    read_data("20_3_Evening.xlsx")

    # Write Output file
    # Don't Touch
    writer = pd.ExcelWriter("Output.xlsx", engine="xlsxwriter")
    df = pd.DataFrame(data)
    df = df.transpose()
    df.to_excel(writer, sheet_name="Fuck Off")

    workbook = writer.book
    worksheet = writer.sheets["Fuck Off"]

    writer.save()
    
    # Debug if cannot finish
    print("Finished");


main()
