import pylightxl as xl

# TODO: Make the fiel usable for all excel files.

if __name__ == "__main__":
    fn_paper = "Special issues - call for papers.xlsx"
    fn_paper_sheet = "Special issues"
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
    db.add_nr(name="special_issues", ws=fn_paper_sheet, address=f"{table_start_cell}:{table_end_cell}")
    # print(db.nr(name="special_issues"))

    with open("out/special_issue.js", "w") as f:
        f.write("var dataSet = [\n")
        for line in db.nr(name="special_issues")[1:]:
            f.write("[\n '")
            f.write("', '".join(x.strip().replace("\n", " ") for x in line[:-3]))
            f.write("', '")
            f.write(f'<a href="{line[-3]}">Link</a>')
            f.write("', '")
            f.write("', '".join(x.strip().replace("\n", " ") for x in line[-2:]))
            f.write("' ],\n")
        f.write("];")
        