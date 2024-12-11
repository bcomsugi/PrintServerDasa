import xlwings as xw
import win32print
import json
import math
import configparser

config = configparser.ConfigParser()
config.read("printserver.ini")
config_sections = config.sections()
print(config.sections())
if 'Selection Printer' in config_sections:
    config_selection_printer = config['Selection Printer']
    activePrinter = config_selection_printer.get('Active_Printer', "Microsoft Print to PDF")
    print(f'{activePrinter = }')

def get_available_printer_names():
    printer_names = []
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    
    printers = win32print.EnumPrinters(flags)
    # print(f'{printers = }')
    for printer in printers:
        # print(printer, type(printer))

        # printer_name = printer['PrinterName']
        _,_,printer_name,_ =  printer
        printer_names.append(printer_name)
    
    return printer_names

def printToPrinter(dt:dict, activePrinter):
    filename = "Packing list-023-2.xlsx"
    filename = "template_packinglist.xlsx"
    filename = "packinglist.xlsx"
    wb = xw.Book(filename)
    sheetNames = wb.sheet_names
    # sheet = xw.Book(filename).sheets[0]
    sheet1 = wb.sheets('Template')
    print(wb.sheet_names)
    if "print 1" not in sheetNames:
        sheet1.copy(after=sheet1, name="print 1")
    else:
        wb.sheets('print 1').delete()
        print(wb.sheet_names)
        sheet1.copy(after=sheet1, name="print 1")
    print(wb.sheet_names)
    # ws_print = wb.sheets('print 1')

    # lrow = sheet1.range('A' + str(sheet1.cells.last_cell.row)).end('up').row

    # # # sheet = out_wb.sheets.active
    # # used_range_rows = (sheet1.api.UsedRange.Row, 
    # #     sheet1.api.UsedRange.Row + sheet1.api.UsedRange.Rows.Count)
    # # used_range_cols = (sheet1.api.UsedRange.Column, 
    # #     sheet1.api.UsedRange.Column + sheet1.api.UsedRange.Columns.Count)
    # # used_range = xw.Range(*zip(used_range_rows, used_range_cols))
    # # used_range.select()
    # print("cell")
    # ws_print.select()
    # ws_print.range('A1').select()
    # # start_cell = sheet1.api.Range('A2')
    # print('start cell')
    
    # cell = ws_print.api.UsedRange.Find("[customername]", After=ws_print.api.Range('A2'), LookIn=-4163,
    #                                  SearchOrder=2 , SearchDirection=1, MatchCase=False)
    # cell2 = ws_print.api.UsedRange.Find("[date]")
    # if cell:
    #     print('after cell', cell._inner.Address, type(cell._inner.Address))
    #     ws_print.range(cell._inner.Address).select()
    #     cell2 = ws_print.api.UsedRange.Find("[customername]", After=ws_print.api.Range(cell._inner.Address))
    #     ws_print.range(cell2._inner.Address).select()
    # if cell2:
    #     print(cell2)
    
    
    #clear all "print" sheetname
    template_sh = wb.sheets('Template')
    linesPerPage = 11
    start_itemline = 7

    for name in wb.sheet_names:
        if "print" in name:
            wb.sheets(name).delete()

    total_page = math.ceil(len(dt.get('LineAdd'))/linesPerPage)
    print(f'{total_page = }')

    sheet_choice = {}
    lines_choice = {}
    lines = []
    if total_page>0:
        for i in range(total_page):
            template_sh.copy(name=f'print {i+1}')
            print(wb.sheet_names)
            sheet_choice[str(i+1)]=wb.sheets(f'print {i+1}')
    print(sheet_choice)
    
    for idx, line in enumerate(dt.get('LineAdd')):
        page = math.ceil((idx + 1)/linesPerPage)
        if idx%linesPerPage == 0 and idx != 0:
            lines_choice[str(page - 1)]=lines
            lines=[]
        lines.append(line)
    lines_choice[str(page)]=lines
    print(f'{lines_choice = }')

#fill data for loop range
    for i in range(total_page):
        ws_print = sheet_choice.get(str(i + 1))
        ws_print.select()
        lines = lines_choice.get(str(i + 1))
        print(f'{ws_print = } {lines = }')
        if total_page>1:
            ws_print.api.UsedRange.Replace("[pages]", str(total_page))
            ws_print.api.UsedRange.Replace("[page]", str(i + 1))
            print(f"page finish {i + 1}")
        else:
            ws_print.api.UsedRange.Replace("[page] of [pages]", "")

        ws_print.api.UsedRange.Replace("[plno]", dt.get('pklist_ID'))
        ws_print.api.UsedRange.Replace("[date]", f"'{dt.get('TxnDate')}")
        ws_print.api.UsedRange.Replace("[customername]", dt.get('CustomerRef_FullName'))
        if dt.get("BillAddress2", None):
            ws_print.api.UsedRange.Replace("[billaddress2]", dt.get('BillAddress2'))
        else:
            ws_print.api.UsedRange.Replace("[billaddress2]", "")
        if dt.get("BillAddress3", None):
            ws_print.api.UsedRange.Replace("[billaddress3]", dt.get('BillAddress3'))
        else:
            ws_print.api.UsedRange.Replace("[billaddress3]", "")
        if dt.get("BillAddress4", None):
            ws_print.api.UsedRange.Replace("[billaddress4]", dt.get('BillAddress4'))
        else:
            ws_print.api.UsedRange.Replace("[billaddress4]", "")
        if dt.get("DT", None):
            ws_print.api.UsedRange.Replace("[dt]", dt.get('DT'))
        else:
            ws_print.api.UsedRange.Replace("[dt]", "")
        if dt.get("PrintCount", None):
            if dt['PrintCount']>0:
                ws_print.api.UsedRange.Replace("[reprint]", "Re-Print")
                dot = ""
                for x in range(dt['PrintCount']):
                    dot = dot + "."
                ws_print.api.UsedRange.Replace("[count]", dot)
        else:
            ws_print.api.UsedRange.Replace("[reprint]", "")
            ws_print.api.UsedRange.Replace("[count]", "")
        print(f"header finish {i + 1}")
        for idx, line in enumerate(lines):
            ws_print.range(f'A{idx+start_itemline}').value = line.get('ItemRef_FullName',':noItem').split(":")[-1]
            ws_print.range(f'B{idx+start_itemline}').value = line.get('Quantity', 'noQty')
            ws_print.range(f'C{idx+start_itemline}').value = line.get('UOM', 'noUM')
            ws_print.range(f'D{idx+start_itemline}').value = line.get('Rack', 'noRack')
            ws_print.range(f'E{idx+start_itemline}').value = line.get('InLineMemo','noInline')
### Printto any Printer(set ActivePrinter)
    for i in range(total_page):
        ws_print = sheet_choice.get(str(i + 1))
        # ws_print.select()
        print(f'{activePrinter = }')
        if activePrinter == None:
            activePrinter = "Microsoft Print to"
        res = ws_print.range("a1:g20").api.PrintOut(ActivePrinter=activePrinter)
        print(f'{res = }')



    # ws_print.delete()


    # print(sheet1.range("A1").value)
    # sheet1.range("A1").value = "CHINA XINO GROUP CO., LTD2"
    # print(sheet1.range("A1").value)
    # # print(sheet1.range("A1:b3").value)

    # sheet1.range((1,1))
    # # sheet1.range((1,1), (3,3)).api.PrintOut()
    # # print(sheet1.range("a1:i13").api.PrintOut(ActivePrinter="Microsoft Print to PDF"))

    # # sheet1.range("a1:i13").to_pdf()
    # # sheet1.to_pdf()
    # # sheet1.range("NamedRange")

def get_active_printer():   # todo:whatif no active printer. add error checking
    config = configparser.ConfigParser()
    config.read("printserver.ini")
    config_sections = config.sections()
    print(config.sections())
    if 'Selection Printer' in config_sections:
        config_selection_printer = config['Selection Printer']
        activePrinter = config_selection_printer.get('Active_Printer', "Microsoft Print to PDF")
        print(f'{activePrinter = }')
    return activePrinter

if __name__ == "__main__":
    available_printers = get_available_printer_names()
    if available_printers:
        print("Available printers:")
        config['Printer List']={}
        for idx, printer_name in enumerate(available_printers):
            print(f'{idx}: {printer_name}')
            config['Printer List'][str(idx)]=printer_name
        print("x: Exit")
        while 1:
            selectedPrinter = input('Choose Which Printer : ')
            if selectedPrinter.lower()=='x' or selectedPrinter.lower()=='q':
                break
            if selectedPrinter.isdecimal():
                if 0 <= int(selectedPrinter) < len(available_printers):
                    activePrinter = config.get('Printer List', selectedPrinter)
                    print(f"{activePrinter} is Selected as Active Printer")
                    break
        if 'Selection Printer' not in config.sections():
            config['Selection Printer'] = {}
        config['Selection Printer']['Active_Printer'] = activePrinter

        with open('printserver.ini', 'w') as configFile:
            config.write(configFile)

    else:
        print("No printers found.")

    dt = {'pklist_ID': 68, 'CustomerRef_FullName': 'Toko Cahaya Timur', 'TxnDate': '07-09-2024 05:07', 'created_by': 'sugi', 'LineAdd': [{'ItemRef_FullName': 'TACO:W:TH-001AA', 'Quantity': 3, 'InLineMemo': 'Krian'}, {'ItemRef_FullName': 'TACO:G_D:TH-011G', 'Quantity': 1, 'InLineMemo': None}, {'ItemRef_FullName': 'TACO:W:TH-001AA', 'Quantity': 4, 'InLineMemo': None},
                                                                                                                                         {'ItemRef_FullName': 'TACO:W:TH-001AA', 'Quantity': 3, 'InLineMemo': 'Krian'}, {'ItemRef_FullName': 'TACO:G_D:TH-011G', 'Quantity': 1, 'InLineMemo': None}, {'ItemRef_FullName': 'TACO:W:TH-001AA', 'Quantity': 4, 'InLineMemo': None},
                                                                                                                                         {'ItemRef_FullName': 'TACO:W:TH-001AA', 'Quantity': 3, 'InLineMemo': 'Krian'}, {'ItemRef_FullName': 'TACO:G_D:TH-011G', 'Quantity': 1, 'InLineMemo': None}, {'ItemRef_FullName': 'TACO:W:TH-001AA', 'Quantity': 4, 'InLineMemo': None},
                                                                                                                                         {'ItemRef_FullName': 'TACO:W:TH-001AA', 'Quantity': 3, 'InLineMemo': 'Krian'}, {'ItemRef_FullName': 'TACO:G_D:TH-011G', 'Quantity': 1, 'InLineMemo': None}, {'ItemRef_FullName': 'TACO:W:TH-001AA', 'Quantity': 4, 'InLineMemo': None}]}
    # printToPrinter(dt, activePrinter)
