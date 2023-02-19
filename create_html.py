import pylightxl as xl
from datetime import date, datetime
import os.path
import re


def check_expiry_date(given_date):
    """
    Check if the date in the excel file is past from today or not
    :param given_date: date from the excel file
    :return: true or false
    """
    if given_date == "Löpande":
        return False
    today = date.today()
    r = re.compile(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})")
    matches = r.findall(given_date)
    y = int(matches[0][0])
    m = int(matches[0][1])
    d = int(matches[0][2])
    received_date = date(y, m, d)
    if received_date<=today:
        return True
    return False


def generate_out(file_params):
    fn_paper = file_params["file_name"]
    fn_paper_sheet = file_params["sheet_name"]
    table_name = file_params["out_name"]

    # Check if file exists, otherwise skip the file for processing
    if not os.path.isfile(fn_paper):
        return

    print(f"Processing {fn_paper}")
    db = xl.readxl(fn_paper)
    
    # Find the starting row of the table
    # iterates through column B to find the first cell with data
    check_row = 1
    check_col = 2
    value = db.ws(ws=fn_paper_sheet).index(row=check_row, col=check_col)
    while value == '':
        check_row = check_row + 1
        value = db.ws(ws=fn_paper_sheet).index(row=check_row, col=check_col)
    table_start_row = check_row
    table_start_cell = f"A{check_row}"

    # Find end column of table
    check_row = table_start_row
    check_col = 1
    value = db.ws(ws=fn_paper_sheet).index(row=check_row, col=check_col)
    while value != '':
        check_col = check_col + 1
        value = db.ws(ws=fn_paper_sheet).index(row=check_row, col=check_col)
    table_end_col = check_col - 1

    # Find end row of table
    check_row = table_start_row
    check_col = 1

    value = db.ws(ws=fn_paper_sheet).index(row=check_row, col=check_col)
    while value != '':
        check_row = check_row + 1
        value = db.ws(ws=fn_paper_sheet).index(row=check_row, col=check_col)
    table_end_row = check_row - 1
    table_end_cell = xl.pylightxl.utility_index2address(table_end_row, table_end_col)


    # Define the table data
    db.add_nr(name=table_name, ws=fn_paper_sheet, address=f"{table_start_cell}:{table_end_cell}")

    # Find index of Link column to make it hyperlink
    header = db.nr(name=table_name)[0]
    link_idx = header.index("Link")

    # find index of date in the file
    date_col_names = ['Dates', 'Closing date', 'Deadline']
    for date_col_name in date_col_names:
        if date_col_name in header:
            date_idx = header.index(date_col_name)

    with open(f"{table_name}.js", "w") as f:
        f.write("var dataSet = [\n")
        for line in db.nr(name=table_name)[1:]:
            # If the date has expired, do not add it to the list 
            if check_expiry_date(line[date_idx]):
                continue
            
            f.write("[\n '")
            f.write("', '".join(x.strip().replace("\n", " ") for x in line[:link_idx]))
            f.write("', '")
            f.write(f'<a href="{line[link_idx]}" target="_blank">Link</a>')
            f.write("', '")
            f.write("', '".join(x.strip().replace("\n", " ") for x in line[link_idx+1:]))
            f.write("' ],\n")
        
        f.write("];")
        

def file_edit_date():
    print("Adding update date")
    date_time = datetime.now()
    date_time = f"{date_time.year}-{date_time.month}-{date_time.day} {date_time.hour}:{date_time.minute}"
    with open("updated_date.js", "w") as f:
        f.write(f"let updated_date = '{date_time}';")


if __name__ == "__main__":
    file_params = [{"file_name": "Special issues - call for papers.xlsx",
                    "sheet_name": "Special issues",
                    "out_name": "special_issues"},
                    {"file_name": "Seminars webinars etc.xlsx",
                    "sheet_name": "Seminars, WS, etc",
                    "out_name": "seminar_webinar"},
                    {"file_name": "Research conference calls.xlsx",
                    "sheet_name": "Research conf",
                    "out_name": "research_conf"},
                    {"file_name": "Research calls.xlsx",
                    "sheet_name": "Research calls",
                    "out_name": "research_calls"},
                ]
    
    for file_param in file_params:
        generate_out(file_param)

    file_edit_date()
    print("Completed")