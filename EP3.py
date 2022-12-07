import xlrd
import random as rd
import xlsxwriter as xls
import colorama
from colorama import Fore, Back, Style
colorama.init(autoreset=True)

#AllGroups.xlsx dosyasini acar ve islem yapilacak tabloyu secer
file = "AllGroups.xlsx"
allGroupsWorkbook = xlrd.open_workbook(file)
allGroupsSheet = allGroupsWorkbook.sheet_by_index(1) #Excel dosyasindaki 0'inci indexteki sayfa islem yapicagimizi belirtiyoruz.

#satir sayisini
print(f"{Fore.LIGHTYELLOW_EX}Bu excel sayfasinda {allGroupsSheet.nrows} satir bulunmaktadir")
#sutun sayisini
print(f"{Fore.LIGHTYELLOW_EX}Bu excel sayfasinda {allGroupsSheet.ncols} sutun bulunmaktadir")

#Zimmetleme isleminin rapor edilecegi Duzenlenmis.xlsx dosyasini olusturur ve gerekli duzenlemelerle basliklari atar. Gruplardaki uye sayisini da icerir.
duzenlenmisWorkbook = xls.Workbook("Duzenlenmis.xlsx") #Duzenlenmis.xlsx adlı dosyayı olusturur.
duzenlenmisSheet = duzenlenmisWorkbook.add_worksheet("Distibutions Groups") #Excel dosyasina tablo ismi verir
duzenlenmisSheet.write("A1","Group Name")
duzenlenmisSheet.write("B1","Owner")
duzenlenmisSheet.write("C1","Number of Members ")

#Toplam Kisi havuzundakileri listeledigimiz Kisiler.xlsx dosyasi olusturulur.
kisilerWorkbook = xls.Workbook("Kisiler.xlsx")
kisilerSheet= kisilerWorkbook.add_worksheet("Groups Members")
kisilerSheet.write("A1","Member Names")
kisilerSheet.write("B1","Owner of Number")

members_names = [] #Toplam kisi havuzunun tespiti icin olusturulmus liste
members = [] #Bir hucredeki uyelerin gecici tutuldugu liste

for x in range(1,allGroupsSheet.nrows):#1 den baslamamizin sebebi basliklari alsin istemiyoruz
    #İlgili hucrelerdeki kisilerin duzenlenmesi icin on islem
    members.append((allGroupsSheet.cell_value(x, 6)))
    members = "".join(members)  # Array icindeki veriyi string'e donusturur.
    members = members.split(";")  # String'i ";" isaretine gore string listesine donusturur.
    members = list(members)
    list_size = len(members)
    owner = rd.choice(members)
    #Duzenlenmis.xlsx dosyasina grup adlarini,grup sahiplerini ve grupta bulunan uye sayisini yazar
    duzenlenmisSheet.write(f"A{x + 1}", allGroupsSheet.cell_value(x, 1))
    duzenlenmisSheet.write(f"B{x + 1}", owner)
    duzenlenmisSheet.write(f"C{x + 1}", list_size)
    print(f"{allGroupsSheet.cell_value(x,1)} ===> {owner} Kisisine Zimmetlendi")
    members_names.extend(members)
    members = [] #members degiskeni bir alt hucrenin verilerini almasi icin bosaltilir.

duzenlenmisWorkbook.close() #İslem bittikten sonra Duzenlenmis.xlsx dosyasi kapatilir.

members_names = set(members_names)
memberOfNumber = len(members_names)
members_names = list(members_names)
members_names.sort()

for x in range(1,memberOfNumber):
    kisilerSheet.write(f"A{x+1}",members_names[x])
    #kisilerSheet.write(f"B{x+1}","NULL")
kisilerWorkbook.close()


print(f"{Fore.LIGHTYELLOW_EX}Toplam Kişi Havuzu ===> {memberOfNumber}")

