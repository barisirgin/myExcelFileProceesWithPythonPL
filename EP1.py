import xlrd
import colorama
from colorama import Fore, Back, Style
colorama.init(autoreset=True)

file = "filename1.xlsx"
allGroupsWorkbook = xlrd.open_workbook(file)
allGroupsSheet = allGroupsWorkbook.sheet_by_index(0)

file2 = "filename2.xlsx"
allGroupsWorkbook2 = xlrd.open_workbook(file2)
allGroupsSheet2 = allGroupsWorkbook2.sheet_by_index(0)

duzenlenmis1 = []
duzenlenmis2 = []

for x in range(1,allGroupsSheet.nrows):
     duzenlenmis1.append((allGroupsSheet.cell_value(x, 0)))

for y in range(1,allGroupsSheet2.nrows):
     duzenlenmis2.append((allGroupsSheet2.cell_value(y, 0)))

duzenlenmis1 = set(duzenlenmis1)
duzenlenmis2 = set(duzenlenmis2)
degisiklik = duzenlenmis1.difference(duzenlenmis2)
degisiklik2 = duzenlenmis2.difference(duzenlenmis1)

print("\n\n\n\n\n\n")
print(f"                      {Fore.RED}Toplam {len(degisiklik)} grup cikartilmis.")
print(f"{Fore.CYAN} {degisiklik}\n\n")
print(f"                      {Fore.RED}Toplam {len(degisiklik2)} grup eklenmis.")
print(f"{Fore.CYAN} {degisiklik2}\n\n\n")

