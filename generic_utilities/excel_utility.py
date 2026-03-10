import xlrd
xlrd.xlsx.ensure_elementtree_imported(False,None)
xlrd.xlsx.Element_has_iter=True
path = r'C:\selenium_project\files_\demo_testdata.xlsx'


def read_excel():
    workbook = xlrd.open_workbook(path)                     ## book obj
    worksheet = workbook.sheet_by_name("Sheet1")            ## sheet_obj
    rows = worksheet.get_rows()                             ## generator_obj
    header = next(rows)

    d = {}
    for ele in rows:
        d[ele[0].value] =  ele[1].value

    return d














































































