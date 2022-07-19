import pandas as pd
import xlsxwriter
import xlwings as xw
from pandas import DataFrame


def get_schd_B_data():
    # Write the datasheet, CC listing, Pool Costs

    path = r'datasheet.xlsx'

    nicra_df = pd.read_excel(path,
                             sheet_name="Datasheet", index_col=None)

    nicra_df: DataFrame = nicra_df.astype(str)

    # Update cc
    filt_3a = ((nicra_df['Cost_centre'].str.contains('907') == True) | (
            nicra_df['Cost_centre'].str.contains('917') == True) | (
                       nicra_df['Cost_centre'].str.contains('927') == True))

    nicra_df.loc[filt_3a, 'Co Management(T)'] = 'HQ'

    # b. Update Co Management
    filt_3b = (nicra_df['Cost_centre'] == '90258')

    nicra_df.loc[filt_3b, 'Co Management(T)'] = 'SCI CO'

    # c. Update Account Rep based on account text
    filt_3c = (nicra_df['Account(T)'].str.contains('FG-') == True)

    nicra_df.loc[filt_3c, 'Acc_rep(T)'] = 'Grants to other Orgs'

    # 5. Manual Updates SOP 10.25.21
    pool_df = pd.read_excel(path,
                            sheet_name='Pool_Cost')

    # a. Update Builiding Ops
    filt_5a = pool_df['Cost_centre'].isin([90272, 90345, 90347, 91350])

    pool_df.loc[filt_5a, 'Pool_costs'] = 'B'

    # b. Update Support Services
    filt_5b = pool_df['Cost_centre'].isin([90341, 90342, 90343, 90344])

    pool_df.loc[filt_5b, 'Pool_costs'] = 'S'

    # Get cc_listing
    cc_listing_df = pd.read_excel(path,
                                  sheet_name='cc_listing', usecols='A:C')

    return cc_listing_df, nicra_df, pool_df


"""

# Insert user_Input_GUI.py
def call_cc_GUI ():

"""


def write_sched_B_data(cc_listing_df, nicra_df, pool_df):
    cc_listing_np = cc_listing_df.to_numpy()

    def create_wb():
        workbook = xlsxwriter.Workbook('output.xlsx')
        sch_b_ws = workbook.add_worksheet('Sch B-HQ Costs Alloc')
        workbook.add_worksheet('NICRA datasheet')
        workbook.add_worksheet('pool_costs')
        workbook.add_worksheet('cc_listing')
        return workbook, sch_b_ws

    def write_top_section(workbook, sch_b_ws):
        global bold
        global cc_count
        bold = workbook.add_format({'bold': True})
        cc_count = cc_listing_df.shape[0]

        f = open('row_3_headers.txt', 'r')

        def write_top_headers():

            row = 2
            col = 3
            # Excel Header cc_listing ROW 4
            for line in f.readlines():
                sch_b_ws.write(row, col, line, bold)
                col += 1
            f.close()

        write_top_headers()


        def write_top_detail_headers():
            f = open('row_4_headers.txt', 'r')
            row = 2
            col = 3
            # Excel Header cc_listing ROW 4
            for line in f.readlines():
                sch_b_ws.write(row, col, line, bold)
                col += 1
            f.close()

        write_top_detail_headers()

        def write_total_expenses():
            col = 0
            row = 4
            sch_b_ws.write(row, col, "Total Field", bold)
            col += 3
            with open('total_expense_formulas.txt', 'r') as formula_file:
                for formula_line in formula_file:
                    row_num_str = str(row + 1)  # plus one because excel and python use a different index
                    formula_row_variable = "\" + row_as_str + \""  # set new language variable
                    formula = formula_line.replace(formula_row_variable, row_num_str)  #
                    sch_b_ws.write_formula(row, col, formula)
                    col += 1

        write_total_expenses()

        def write_cost_centres():
            # row and column for Schd_B_Loop
            row = 7
            col = 0
            # cc_listing data to Excel and Excel formulas
            for cc_text, cc_number, cc_location in cc_listing_np:
                sch_b_ws.write(row, col, cc_number)
                sch_b_ws.write(row, col + 1, cc_text)
                sch_b_ws.write(row, col + 2, cc_location)
                col += 1

        write_cost_centres()

        def write_sch_b_formulas():
            col = 3

            # Reading from formula file
            # Each line is a column in the xlsx
            with open('formulas.txt', 'r') as formula_file:
                for formula_line in formula_file:
                    # Inital row
                    initial_row = 7
                    final_row = cc_count + initial_row
                    # number of rows determined by a range
                    for row in range(initial_row, final_row):
                        row_num_str = str(row + 1)  # plus one because excel and python use a different index
                        formula_row_variable = "\" + row_as_str + \""  # set new language variable
                        formula = formula_line.replace(formula_row_variable, row_num_str)  #
                        sch_b_ws.write_formula(row, col, formula)
                    col += 1

        write_sch_b_formulas()

        workbook.close()

    write_top_section(*create_wb())

write_sched_B_data(*get_schd_B_data())


def dump_data(cc_listing_df, nicra_df, pool_df, ):
    # load workbook
    app = xw.App(visible=False)
    wb = xw.Book('output.xlsx')
    ws_datasheet = wb.sheets['NICRA datasheet']
    ws_pool = wb.sheets['pool_costs']
    ws_cc = wb.sheets['cc_listing']
    ws_sched_b = wb.sheets['Sch B-HQ Costs Alloc']

    # Update workbook at specified range
    ws_datasheet.range('A1').options(index=False).value = nicra_df
    ws_pool.range('A1').options(index=False).value = pool_df
    ws_cc.range('A1').options(index=False).value = cc_listing_df
    ws_sched_b.range('A7').options(index=False).value = cc_listing_df

    # Close workbook
    wb.save()
    wb.close()
    app.quit()


dump_data(*get_schd_B_data())
