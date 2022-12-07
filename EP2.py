import xlrd
import xlsxwriter as xls
import colorama
from colorama import Fore, Back, Style
colorama.init(autoreset=True)

file = "file1.xlsx"
allGroupsWorkbook = xlrd.open_workbook(file)
allGroupsSheet = allGroupsWorkbook.sheet_by_index(0)

file2 = "file2.xlsx"
allGroupsWorkbook2 = xlrd.open_workbook(file2)
allGroupsSheet2 = allGroupsWorkbook2.sheet_by_index(0)

duzenlenmis1 = []
duzenlenmis2 = []

for x in range(1,allGroupsSheet.nrows):
     duzenlenmis1.append((allGroupsSheet.cell_value(x, 1)).replace(' ','.').lower())
     
print(f"{Fore.RED}{file} personellerin toplam sayisi :{len(duzenlenmis1)}")
duzenlenmis1.sort()

for i in duzenlenmis1:
     print(i)

for z in range(1,allGroupsSheet2.nrows):
     duzenlenmis2.append((allGroupsSheet2.cell_value(z, 2)).replace(' ','.').lower())

"""""
for q in range(4,allGroupsSheet2.nrows):
     duzenlenmis2.append((allGroupsSheet2.cell_value(q, 5)).replace(' ','.'))
"""""

print(f"{Fore.RED}{file2} personellerin toplam sayisi :{len(duzenlenmis2)}")
duzenlenmis2.sort()

for u in duzenlenmis2:
     print(u)

duzenlenmis1 = set(duzenlenmis1)
duzenlenmis2 = set(duzenlenmis2)

degisiklik = duzenlenmis1.difference(duzenlenmis2)
degisiklik2 = duzenlenmis2.difference(duzenlenmis1)
#degisiklik2.remove("")
degisiklik3 = duzenlenmis1.intersection(duzenlenmis2)

t = len(degisiklik)
g = len(degisiklik2)
b = len(degisiklik3)

print("\n\n\n")
print(f"                      {Fore.RED}Toplam {t} kisi file1 listesinde olup file2 listesinde bulunmuyor.")
print(f"{Fore.CYAN} {degisiklik}\n\n")
print(f"                      {Fore.RED}Toplam {g} kisi file2 listesinde olup file1 listesinde bulunmuyor.")
print(f"{Fore.CYAN} {degisiklik2}\n\n\n")
print(f"                      {Fore.RED}Toplam {b} kisi iki listenin kesisimi.")
print(f"{Fore.CYAN} {degisiklik3}\n\n\n")

kumelenmisWorkbook = xls.Workbook("Kiyaslama_file1-file2.xlsx")
kumelenmisSheet = kumelenmisWorkbook.add_worksheet("file1-file2") #Excel dosyasina tablo ismi verir
kumelenmisSheet2 = kumelenmisWorkbook.add_worksheet("file2-file1") #Excel dosyasina tablo ismi verir
kumelenmisSheet3 = kumelenmisWorkbook.add_worksheet("Ortak Kisi Havuzu") #Excel dosyasina tablo ismi verir
kumelenmisSheet.write("A1","file1 listesinde olup file2 listesinde bulunmayanlar")
kumelenmisSheet2.write("A1","file2 listesinde olup file1 listesinde bulunmayanlar")
kumelenmisSheet3.write("A1","Her iki listede ortak olanlar ")

degisilik11 = list(degisiklik)
degisilik12 = list(degisiklik2)
degisilik13 = list(degisiklik3)

for x in range(1,t+1):
    kumelenmisSheet.write(f"A{x + 1}", degisilik11[x-1])
for y in range(1,g+1):
    kumelenmisSheet2.write(f"A{y + 1}", degisilik12[y-1])
for z in range(1,b+1):
    kumelenmisSheet3.write(f"A{z + 1}", degisilik13[z-1])

kumelenmisWorkbook.close()

input()