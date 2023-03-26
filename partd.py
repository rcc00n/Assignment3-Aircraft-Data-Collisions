from openpyxl import load_workbook
from openpyxl.chart import Reference, BarChart

workbook = load_workbook("aircraftWildlifeStrikes.xlsx")
sheet2 = workbook.create_sheet(title="ChartForAirlines")

firstSheet = workbook[workbook.sheetnames[0]]


def get_airline_cells(cells: list) -> list:
    for column_values in firstSheet:
        cells.append(column_values[5].value)
    return cells


def get_cell_values(airlines: list) -> list:
    for airline in firstSheet:
        airlines.append(airline[5].coordinate)
    return airlines


def get_cells_and_airlines(incident_airlines: list, airline_cells: list) -> list:
    return list(
        zip(get_cell_values(incident_airlines), get_airline_cells(airline_cells)))


def get_count_of_airlines(list_cells_airlines: list) -> list:
    airlines = [list_cells_airlines[i][1]
                for i in range(len(list_cells_airlines))]

    # return set(airlines)
    return [[airline, airlines.count(airline)] for airline in set(airlines)]


airline_cells = []
incident_airlines = []
rows = get_count_of_airlines(
    get_cells_and_airlines(incident_airlines, airline_cells))
rows.remove(['Operator', 1])

# credit for code below - https://stackoverflow.com/questions/65679123/sort-nested-list-data-in-python
rows = sorted(rows, key=lambda x: x[0])
print(rows)
print("\n\n\n")
print(len(rows))
print("\n\n\n")

max_collisions = max(rows, key=lambda x: x[1])
print(max_collisions)
# print(type(max_collisions))

len_of_rows = len(rows)

row_no = 0

while row_no < len_of_rows:
    if 'UNKNOWN' in rows[row_no]:
        rows.pop(row_no)
        len_of_rows -= 1
        row_no -= 1
    elif rows[row_no][1] < (max_collisions[1]*0.1):
        rows.pop(row_no)
        len_of_rows -= 1
        row_no -= 1
    row_no += 1


print(rows)


for row in rows:
    sheet2.append(row)

chart1 = BarChart()
chart1.type = "col"
chart1.style = 10
chart1.title = "Airlines and # of Collisions"
chart1.y_axis.title = 'Frequency'
chart1.x_axis.title = 'Airlines'

data = Reference(sheet2, min_col=2, min_row=1, max_row=len(rows), max_col=2)
cats = Reference(sheet2, min_col=1, min_row=1, max_row=len(rows), max_col=1)
chart1.add_data(data)
chart1.set_categories(cats)
chart1.legend = None
sheet2.add_chart(chart1, "D2")
workbook.save('aircraftWildlifeStrikes.xlsx')
