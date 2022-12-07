# -*- coding: utf-8 -*-

from email import message
import win32com.client as client
import xlrd
#import xlsxwriter
import colorama
import time as t
from colorama import Fore, Back, Style
from datetime import date
colorama.init(autoreset=True)
import datetime
from dateutil.relativedelta import relativedelta
import time as tc
import os

# Excelden cekilen tarihler 01/01/1900 tarihinden günümüze gecen gun sayisi olarak geldigi icin bu degiskeni kullanmamiz gerekecek
excel_baslangic_tarihi = datetime.date(1900,1,1) 

#Internet kontrolu
while True:

    hostname = "google.com"
    #-n windwos -c linux için
    response = os.system("ping -n 1 " + hostname)
    t.sleep(2)
    if response ==0:
        print(f"{hostname} is up!")
        t.sleep(2)
        break
    else:
        print(f"{hostname} is down!")
        t.sleep(2)

# Excel dosyasinin adini input ile kullanicidan hata yapma ihtimalinide denetleyen kosullu yapi ile python ortaminda acilmasi
"""""
#İslem yapacagimiz xlsx dosyasini acar ve islem yapilacak tabloyu seccer
while True:
    file = input("Excel dosya adi : ")
    file2 = str(file) + ".xlsx"
    try: 
        Workbook1 = xlrd.open_workbook(file2)
        Sheet1 = Workbook1.sheet_by_index(1) #Excel dosyasindaki 1'inci indexteki sayfa islem yapicagimizi belirtiyoruz.
        break
    except FileNotFoundError:
        print(f"{Fore.RED}! ! ! Girdiginiz dosya adi hatali veya uygulama ile ayni dizinde bulunmuyor.Tekrar deneyin.")
"""""        

# Sabit excel dosyasinin python ortaminda acilmasi 
file = "fileName.xlsx"
Workbook1 = xlrd.open_workbook(file)
Sheet1 = Workbook1.sheet_by_index(1) #Excel dosyasindaki 1'inci indexteki sayfa (2.sayfa) ile islem yapicagimizi belirtiyoruz.


# Tablodan cekilen ve isimize yarayacak verilerin listeye aktarildigi bolum
danisman = [] # danismanlarin isimlerinin tutuldugu liste
sorumlu = [] # sorumlularin isimlerinin tutuldugu liste
tarih = [] # date türünde tarihin tutuldugu ve bugüne gore kac gun kaldigini hesaplatmak icin tarihleri tuttugumuz liste
tarih_date = [] # mailde belirtmek icin string türünde tarih tutulan liste
kalangun_list = [] # bugüne gore kalan gunlerin tutuldugu liste
mail_adress = [] # exceldeki veriler ile sormlularin mail adreslerinin tutulacagi liste

# Tablodaki gerekli verileri listeye atar.
for x in range(4,Sheet1.nrows):
    danisman.append(Sheet1.cell_value(x, 4))
    sorumlu.append(Sheet1.cell_value(x, 8))
    tarih.append(relativedelta(days =(int(Sheet1.cell_value(x, 10)))-2) + excel_baslangic_tarihi)
    tarih_date.append(str(relativedelta(days =(int(Sheet1.cell_value(x, 10)))-2) + excel_baslangic_tarihi))
    mail_adress.append(f"{str(Sheet1.cell_value(x, 8).lower()).replace(' ','.')}@{Sheet1.cell_value(x, 7)};")

# Nesne sayısını degisken olarak tutar    
islem_sayisi = len(sorumlu)    
for y in range(len(tarih)):
       kalangun_list.append((tarih[y] - date.today()).days)

mail_text = []
for z in range(islem_sayisi):
    mail_text.append(f"{sorumlu[z]} kisisine bagli {danisman[z]} kisisinin danismanlik süresinin sonlanmasina {kalangun_list[z]} gun kalmistir. {tarih_date[z]} tarihinde hesabi kapatilacaktir.")
    print(f"{z+1} - {sorumlu[z]} kisisine bagli {danisman[z]} kisisinin danismanlik süresinin sonlanmasina {kalangun_list[z]} gun kalmistir. {tarih_date[z]} tarihinde hesabi kapatilacaktir.----> Sorumlunun mail adresi [{mail_adress[z]}]")

danisman2 = danisman.copy()
sorumlu2 = sorumlu.copy()
tarih2 = tarih.copy()
tarih_date2 = tarih_date.copy()
kalangun_list2 = kalangun_list.copy()
mail_adress2 =  mail_adress.copy()
mail_text2 = mail_text.copy()

for q in range(islem_sayisi):
    #if kalangun_list[q] != 10: 
    if kalangun_list[q] > 10:
        danisman2.remove(danisman[q])
        sorumlu2.remove(sorumlu[q])
        tarih2.remove(tarih[q])
        tarih_date2.remove(tarih_date[q])
        mail_adress2.remove(mail_adress[q])
        mail_text2.remove(mail_text[q])
        kalangun_list2.remove(kalangun_list[q])
    else :
        continue

mail_atilacak_sayi = len(kalangun_list2)

print(f"{Fore.RESET}------------------------------------------------------------------------------------------------------------------------------------------------------\n\n\n\n")
for j in range(mail_atilacak_sayi):
    print(f"{Fore.YELLOW}{j+1} - {sorumlu2[j]} kisisine bagli {danisman2[j]} kisisinin danismanlik süresinin sonlanmasina {kalangun_list2[j]} gun kalmistir. {tarih_date2[j]} tarihinde hesabi kapatilacaktir.----> Sorumlunun mail adresi [{mail_adress2[j]}]")

if mail_atilacak_sayi != 0:
    #Mail atma kod blogu
    for i in range(mail_atilacak_sayi):
        outlook = client.Dispatch("Outlook.Application")
        message = outlook.CreateItem(0)
        message.Display()

        message.To = mail_adress2[i]
        #message.CC = "bt.dijitalisyeri@logo.com.tr" 

        message.Subject = "Danışman Süre Sonu Bildirimi"
        message.Body = f"Merhaba,\n\nSize bağlı olan danışmanlardan aşağıda ilettiğim süre sonu yaklaşan kişi için;\n\nDanışman Adı :  {danisman2[i]}\nHesap Kapama Tarihi : {tarih_date2[i]}\n\n1-Süre uzatımı olacak mıdır? Olacak ise [https://link] linkinden süre uzatımı için kayıt açabilir misiniz?\n\n2-Olmayacak ise hesabı disable etmemiz gerekecek.\n\n\nBilginize.\n\nİyi Çalışmalar."
        message.Save()
        t.sleep(0.01)
        message.Send()

    print(f"{Fore.BLUE}İşlem Bitti. Toplam {mail_atilacak_sayi} tane mail gonderildi. Uygulamayi kapatmak icin Enter tusuna basiniz...")
    input()
else:
    print(f"{Fore.BLUE}Son 10 gün kalan hesap yok bu yüzden mail atilmadi...")
    print(f"{Fore.BLUE}Uygulamayi kapatmak icin Enter tusuna basiniz...")

if os.path.exists("Danışman Süre Sonu Bildirimi.xlsx"):
  os.remove("Danışman Süre Sonu Bildirimi.xlsx")
  print(f"{Fore.RED}Dosya silindi...")
else:
  print(f"{Fore.RED}Dosya mevcut değil...")

input()


#Listelerin elemanlarinin Konsolda goruntulendigi bolum
"""""
print(f"{Fore.BLUE}DANISMANLAR")
for danis in danisman: 
    print(f"{danis}")
print(f"{Fore.RED}Danisman sayisi : {len(danisman)}\n")

print(f"{Fore.BLUE}SORUMLULAR")
for sorum in sorumlu:
     print(f"{sorum}")
print(f"{Fore.RED}Sorumlu sayisi : {len(sorumlu)}\n")

print(f"{Fore.BLUE}TARIH Turu")
for tar in tarih:
    print(tar)
print(f"{Fore.RED}Tarih sayisi : {len(tarih)}\n")

print(f"{Fore.BLUE}TARIH")
for tur in tarih_date:
    print(tur)
print(f"{Fore.RED}StringTarih sayisi : {len(tarih_date)}\n")

print(f"{Fore.BLUE}KALAN GUN")
for kalan in kalangun_list:
    print(kalan)
print(f"{Fore.RED}Kalangun liste eleman sayisi : {len(kalangun_list)}\n")
"""""
