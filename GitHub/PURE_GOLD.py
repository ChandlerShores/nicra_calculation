import xlsxwriter

with xlsxwriter.Workbook('output.xlsx') as workbook:
    worksheet = workbook.add_worksheet('Sheet 1')

    # Initial column
    col = 3

    # Reading from formula file
    # Each line is a column in the xlsx
    with open('formulas.txt', 'r') as formula_file:
        for formula_line in formula_file:
            # Inital row
            row = 0
            # number of rows determined by a range
            for row in range(1, 50):
                row_num_str = str(row + 1) # plus one because excel and python use a different index
                formula_row_variable = "\" + row_as_str + \"" # set new language variable
                formula = formula_line.replace(formula_row_variable , row_num_str) #
                worksheet.write_formula(row, col, formula)
            col += 1
