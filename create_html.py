import pylightxl as xl

# TODO: Make the file usable for all excel files.

def generate_out(file_params):
    fn_paper = file_params["file_name"]
    fn_paper_sheet = file_params["sheet_name"]
    table_name = file_params["out_name"]

    db = xl.readxl(fn_paper)
    start = db.ws(ws=fn_paper_sheet).address(address="A5")
    print(start)
    start = db.ws(ws=fn_paper_sheet).address(address="B2")
    print(start)
    
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
    print(table_start_cell)

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
    print(table_end_cell)


    # Define the table data
    db.add_nr(name=table_name, ws=fn_paper_sheet, address=f"{table_start_cell}:{table_end_cell}")
    # print(db.nr(name="special_issues"))

    # Find index of Link column
    header = db.nr(name=table_name)[0]
    print(header)
    link_idx = header.index("Link")
    print(link_idx)


    with open(f"out/{table_name}.js", "w") as f:
        f.write("var dataSet = [\n")
        for line in db.nr(name=table_name)[1:]:
            f.write("[\n '")
            f.write("', '".join(x.strip().replace("\n", " ") for x in line[:link_idx]))
            f.write("', '")
            f.write(f'<a href="{line[link_idx]}" target="_blank">Link</a>')
            f.write("', '")
            f.write("', '".join(x.strip().replace("\n", " ") for x in line[link_idx+1:]))
            f.write("' ],\n")
        f.write("];")
        

if __name__ == "__main__":
    file_params = [{"file_name": "in/Special issues - call for papers.xlsx",
                    "sheet_name": "Special issues",
                    "out_name": "special_issues"},
                    {"file_name": "in/Seminars webinars etc.xlsx",
                    "sheet_name": "Seminars, WS, etc",
                    "out_name": "seminar_webinar"},
                    {"file_name": "in/Research conference calls.xlsx",
                    "sheet_name": "Research conf",
                    "out_name": "research_conf"},
                    {"file_name": "in/Research calls.xlsx",
                    "sheet_name": "Research calls",
                    "out_name": "research_calls"},
                ]

    for file_param in file_params:
        generate_out(file_param)
