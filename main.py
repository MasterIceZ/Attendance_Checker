import pandas as pd


debug_mode = False

def main():
    all_students = pd.read_excel('students.xlsx', sheet_name="Students")

    # Must be update in 1.2
    # print(all_students.head)
    # date = input('date to day : ')
    # interval = input('[M]orning/[A]fternoon/[E]vening : ')
    # table_name = date + " " + interval
    # print(table_name)
    
    data_set = ""
    if debug_mode:
        data_set = "testset.xlsx"
    else:
        data_set = input("Enter Data File's name : ")

    attend = pd.read_excel(dataset, sheet_name="Attendance")
    # print(attend.head)

    students = dict()
    name_map = dict()

    for name in all_students['Name']:
        students[name] = False
    
    all_names = []

    for i in range(len(all_students['Name'])):
        current_user = all_students['Username'][i]
        current_name = all_students['Name'][i]
        name_map[current_user] = current_name
        all_names.append(current_name)

    # print(name_map)
    # print(students.keys())

    names = []
    at = []

    for name in attend['Attendee']:
#        print("Name : " + name)
        if name not in name_map.keys():
            continue
        real_name = name_map[name]
        print(real_name)
        students[real_name] = True
        at.append(True)
        names.append(real_name)

    print(students)
        
    it = 0

    while(len(names) != len(students)):
        while(all_names[it] in names):
            it += 1
        names.append(all_names[it])
        at.append(False)

    data = {}
    for i in range(len(students)) :
        data[i] = {"Name": names[i], "Attend": at[i]}
 
    writer = pd.ExcelWriter("Output.xlsx", engine="xlsxwriter")
    df = pd.DataFrame(data)
    df = df.transpose()
    df.to_excel(writer, sheet_name="Fuck Off")

    workbook = writer.book
    worksheet = writer.sheets["Fuck Off"]

    writer.save()
    print("Finished");


main()
