def judge(sr):
    import openpyxl
    file_sr = sr
    file_sc = "Calculate_"+str(file_sr)

    wb = openpyxl.load_workbook(file_sr)
    ws = wb.active
    i = ws.max_column
    print(i)
    return i