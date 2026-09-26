#!/usr/bin/env python3
"""
Writes the .xlsx files the importer's specs read, with LibreOffice.

A reader tested only against files its own author wrote is tested
against the author's idea of the format. These come from a real
producer instead: LibreOffice builds each workbook through its own API
and saves it with its own Excel filter, so what the specs read is what
a person who exported from LibreOffice would hand this application.

    python3 scripts/xlsx-fixtures.py

Needs LibreOffice and its Python bridge (`import uno`). The files are
checked in, so nobody else needs either unless the fixtures change.
"""

import os
import subprocess
import tempfile
import time

import uno
from com.sun.star.beans import PropertyValue

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, '..', 'src', 'sheet', 'fixtures')


def prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def connect():
    profile = tempfile.mkdtemp(prefix='gessosheet-lo-')
    office = subprocess.Popen([
        'soffice', '--headless', '--invisible', '--nologo', '--norestore',
        f'-env:UserInstallation=file://{profile}',
        '--accept=socket,host=127.0.0.1,port=2083;urp;'
    ])
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
    for _ in range(60):
        try:
            context = resolver.resolve('uno:socket,host=127.0.0.1,port=2083;urp;StarOffice.ComponentContext')
            return office, context.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', context)
        except Exception:
            time.sleep(0.5)
    raise RuntimeError('LibreOffice never answered.')


def number_format(document, code):
    formats = document.NumberFormats
    locale = uno.createUnoStruct('com.sun.star.lang.Locale')
    key = formats.queryKey(code, locale, False)
    return key if key != -1 else formats.addNew(code, locale)


def orders(desktop):
    document = desktop.loadComponentFromURL('private:factory/scalc', '_blank', 0, (prop('Hidden', True),))
    sheets = document.Sheets
    sheets.getByIndex(0).Name = 'Orders'
    sheets.insertNewByName('Rates', 1)
    orders = sheets.getByName('Orders')
    rates = sheets.getByName('Rates')

    rates.getCellByPosition(0, 0).String = 'Tax'
    rates.getCellByPosition(1, 0).Value = 0.08

    title = orders.getCellByPosition(0, 0)
    title.String = 'Quarter orders'
    title.CharWeight = 150  # bold
    title.CellBackColor = 0x1F3864
    title.CharColor = 0xFFFFFF
    orders.getCellRangeByName('A1:E1').merge(True)

    for column, heading in enumerate(['Item', 'Units', 'Price', 'Total', 'Due', 'Code', 'Paid']):
        cell = orders.getCellByPosition(column, 2)
        cell.String = heading
        cell.CharWeight = 150
        cell.CellBackColor = 0xD9D9D9

    rows = [('Widget', 3, 4.5, '2026-09-24', '007', True),
            ('Gadget', 5, 6, '2026-10-01', '012', False),
            ('Gizmo', 7, 8.25, '2026-10-15', '100', True),
            ('Doohickey', 2, 19.99, '2026-11-02', '042', False)]
    for offset, (item, units, price, due, code, paid) in enumerate(rows):
        row = 3 + offset
        orders.getCellByPosition(0, row).String = item
        orders.getCellByPosition(1, row).Value = units
        orders.getCellByPosition(2, row).Value = price
        orders.getCellByPosition(3, row).Formula = f'=B{row + 1}*C{row + 1}'
        orders.getCellByPosition(4, row).Formula = f'=DATEVALUE("{due}")'
        orders.getCellByPosition(5, row).String = code
        orders.getCellByPosition(6, row).Formula = '=TRUE()' if paid else '=FALSE()'

    orders.getCellByPosition(0, 7).String = 'Total'
    orders.getCellByPosition(0, 7).CharWeight = 150
    orders.getCellByPosition(3, 7).Formula = '=SUM(D4:D7)'
    orders.getCellByPosition(0, 8).String = 'Tax'
    orders.getCellByPosition(3, 8).Formula = '=D8*TaxRate'
    orders.getCellByPosition(0, 9).String = 'Share of the first'
    orders.getCellByPosition(3, 9).Formula = '=D4/D8'
    orders.getCellByPosition(0, 10).String = 'Tax, read across'
    orders.getCellByPosition(3, 10).Formula = '=D8*Rates.B1'
    orders.getCellByPosition(0, 11).String = 'hidden working'
    orders.Rows.getByIndex(11).IsVisible = False

    currency = number_format(document, '[$$-409]#,##0.00')
    for name in ['C4:C7', 'D4:D9', 'D11']:
        orders.getCellRangeByName(name).NumberFormat = currency
    orders.getCellRangeByName('E4:E7').NumberFormat = number_format(document, 'YYYY-MM-DD')
    orders.getCellRangeByName('D10').NumberFormat = number_format(document, '0.0%')

    orders.Columns.getByIndex(0).Width = 4200  # hundredths of a millimetre
    orders.Columns.getByIndex(4).Width = 2600

    document.NamedRanges.addNewByName('TaxRate', '$Rates.$B$1', orders.getCellByPosition(0, 0).CellAddress, 0)

    controller = document.CurrentController
    controller.setActiveSheet(orders)
    controller.freezeAtPosition(0, 3)

    path = os.path.abspath(os.path.join(OUT, 'orders.xlsx'))
    document.storeToURL(uno.systemPathToFileUrl(path), (prop('FilterName', 'Calc MS Excel 2007 XML'),))
    document.close(True)
    print('wrote', path)


def main():
    os.makedirs(OUT, exist_ok=True)
    office, desktop = connect()
    try:
        orders(desktop)
    finally:
        try:
            desktop.terminate()
        except Exception:
            pass
        office.wait(timeout=30)


if __name__ == '__main__':
    main()
