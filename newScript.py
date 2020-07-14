import arcpy
import xlwt
import xlrd
import csv

# db_path = r"Connection to dbm363.sde"

db_path = r'Y:\Tde\GIS\Dev\Gold Source\Web\Database Connections\MODEL_DIRECT_CONNECTION_VGISSERV.sde'
customer = db_path + r"\MOB.FUSION_CUST_LOC\MOB.FCL"
# tower = db_path + r"\MOB.SECTORS\MOB.LTE_TDD_2500"

fields = ['SITE', 'CONCAT_CUSTOMER_NAME', 'SIGNAL_STRENGTH']

book = xlwt.Workbook()
sheet1 = book.add_sheet("Result")

sheet1.write(0,0, "SITE")
sheet1.write(0,1, "CUSTOMER NAME")
sheet1.write(0,2, "SIGNAL STRENGTH")


with arcpy.da.SearchCursor(customer, fields) as cursor:
    row = 1
    col = 0
    for customer in cursor:
        if int(customer[2]) != 'None':
            sheet1.write(row, col, str(customer[0]))
            sheet1.write(row, col + 1, str(customer[1]))
            sheet1.write(row, col + 2, str(customer[2]))
            row = row + 1
            col = 0
        else:
            row = row
            col = 0


book.save("Result.xls")



