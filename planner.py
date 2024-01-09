import xlsxwriter
import calendar


cal = calendar.Calendar()
cal.setfirstweekday(6)

months_dict = {1: 'January', 2: 'February', 3: 'March', 4: 'April', 5: 'May', 6: 'June', 7: 'July', 8: 'August', 9: 'September', 10: 'October', 11: 'November', 12: 'December'}
columns_dict = {0: 'A', 1: 'C', 2: 'E', 3: 'G', 4: 'I', 5: 'K', 6: 'M'}
col_dict = {0: 'B', 1: 'D', 2: 'F', 3: 'H', 4: 'J', 5: 'L', 6: 'N'}

start_date = '01/2024'
end_date = '06/2024'

start = start_date.split("/")
start_month = int(start[0])
start_year = int(start[1])
end = end_date.split("/")
end_month = int(end[0])
end_year = int(end[1])

workbook = xlsxwriter.Workbook('AcademicPlanner' + '_' + str(start_year) + '-' + str(end_year) + '.xlsx')

months = []
current = [start_month,start_year]
while current[0] != end_month:
    months.append([current[0],current[1]])
    if current[0] == 12:
        current[0] = 1
        current[1] += 1
    else:
        current[0] += 1
months.append([end_month, end_year])

num_months = 0
for i in months: num_months += 1


# Styles for formatting
column = workbook.add_format({ "right": 7,})
border_right = workbook.add_format({ "right": 1,})
border_top = workbook.add_format({ "top": 1,})
border_bottom = workbook.add_format({"bottom": 7,})
month_format = workbook.add_format({ "bold": 1, "border": 1, "align": "center", "valign": "vcenter", "fg_color": "#9fc5e8", "font_size": 20,})
day_header = workbook.add_format({ "bold": 1, "border": 1, "align": "center", "valign": "vcenter", "fg_color": "#3d85c6", "font_size": 15,})
valid_weekend = workbook.add_format({ "bold": 1, "border": 1, "align": "center", "valign": "vcenter", "fg_color": "#6fa8dc", "font_size": 15,})
invalid_weekend = workbook.add_format({ "bold": 1, "border": 1, "align": "center", "valign": "vcenter", "fg_color": "#858585", "font_size": 15,})
valid_weekday = workbook.add_format({"bold": 1, "border": 1, "align": "center", "valign": "vcenter", "fg_color": "#9fc5e8", "font_size": 15,})
invalid_weekday = workbook.add_format({"bold": 1, "border": 1, "align": "center", "valign": "vcenter", "fg_color": "#9a9a9c", "font_size": 15,})
invalid = workbook.add_format({"fg_color": "#cccccc", "border": 7,})

info = workbook.add_worksheet('Class Info + Monthly Notes')
for i in range(0, num_months*2, 2):
    info.set_column(i, i, 3)
for i in range(1, num_months*2+1, 2):
    info.set_column(i, i, 30)
    for a in range(11):
        info.write(a, i, ' ', border_bottom)
for i in range(num_months):
    if i % 2 == 0: info.write(col_dict[i]+str(1), months_dict[months[i][0]], valid_weekend)
    else: info.write(col_dict[i]+str(1), months_dict[months[i][0]], valid_weekday)

info.merge_range("B14:J14", 'Class Info', month_format)
for i in range(1,9):
    info.merge_range('B'+str(14+i)+':J'+str(14+i), ' ', border_bottom)

def create_month(month, year):
    '''Creates sheet for the specified calendar month with proper headings
    and dropdown selection for assignment progress'''

    month_str = months_dict[month]
    month_num = month
    month = workbook.add_worksheet(month_str)

    for i in range(0, 12, 2):
        month.set_column(i, i, 8)
    for i in range(1, 14, 2):
        month.set_column(i, i, 25)
    month.merge_range("A1:N1", month_str, month_format)
    month.merge_range("A2:B2", 'Sunday', day_header)
    month.merge_range("C2:D2", 'Monday', day_header)
    month.merge_range("E2:F2", 'Tuesday', day_header)
    month.merge_range("G2:H2", 'Wednesday', day_header)
    month.merge_range("I2:J2", 'Thursday', day_header)
    month.merge_range("K2:L2", 'Friday', day_header)
    month.merge_range("M2:N2", 'Saturday', day_header)

    weeks = cal.monthdayscalendar(year,month_num)
    i= 3

    prev_month = 0
    prev_year = year
    if month_num == 1:
        prev_month = 12
        prev_year = year - 1
    else: prev_month = month_num - 1
    prev_weeks = cal.monthdayscalendar(prev_year, prev_month)
    last_prev = prev_weeks[-1]

    next_month = 0
    next_year = year
    if month_num == 12:
        next_month = 1
        next_year = year + 1
    else: next_month = month_num + 1
    next_weeks = cal.monthdayscalendar(next_year, next_month)
    first_next = next_weeks[0]

    for week in weeks:
        for index in range(7):
            b = i
            if week[index] == 0:
                if week == weeks[0]:
                    month.merge_range(columns_dict[index]+str(i)+':'+col_dict[index]+str(i), last_prev[index], invalid_weekend)
                else: month.merge_range(columns_dict[index]+str(i)+':'+col_dict[index]+str(i), first_next[index], invalid_weekend)
                for a in range(9):
                    b += 1
                    month.data_validation(columns_dict[index]+str(b), {'validate': 'list', 'source': '=sources!$A$1:$A$3'})
                    month.write_column(columns_dict[index]+str(b), ' ', invalid)
                    month.write_column(col_dict[index]+str(b), ' ', invalid)
                    a + 1
            else: 
                month.merge_range(columns_dict[index]+str(i)+':'+col_dict[index]+str(i), week[index], valid_weekend)
                for a in range(9):
                    b += 1 
                    month.data_validation(columns_dict[index]+str(b), {'validate': 'list', 'source': '=sources!$A$1:$A$3'})
                    month.write_column(columns_dict[index]+str(b), ' ', column)
                    month.write_column(col_dict[index]+str(b), ' ', column)
                    a + 1
        i += 9
    
    month.set_row(i, 12, border_top)
    month.set_row(0, 40)
    month.set_row(1, 20)


for month in months:
    create_month(month[0], month[1])

sources = workbook.add_worksheet('sources')
sources.write('A1', 'Pending')
sources.write('A2', 'WIP')
sources.write('A3', '✅')

workbook.close()