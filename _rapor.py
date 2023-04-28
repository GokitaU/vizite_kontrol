# -*- coding: utf-8 -*-

# Web kazıma işlemleri ve raporlamalar / main dosyası

from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog, QMessageBox, QTableWidgetItem
from PyQt5.QtCore import Qt

from _raporForm import Ui_MainRaporKontrol
from _fileData import DosyaIslem

from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.support.ui import Select # WebDriverWait
# from selenium.webdriver.support import expected_conditions
# from selenium.webdriver.chrome.service import Service
from selenium import webdriver
from selenium.webdriver.common.by import By
# from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.chrome.options import Options
# import chromedriver_autoinstaller


# from tqdm import tqdm
from time import sleep, strftime
from locale import setlocale, LC_ALL
from cv2 import imread
from pandas import DataFrame
from pathlib import Path
from bs4 import BeautifulSoup
from sys import argv, exit
from os import path

import pytesseract
import re

setlocale(LC_ALL, 'turkish')

class RaporKontrol(QMainWindow):
        
    def __init__(self):
        super(RaporKontrol, self).__init__()

        self.ui = Ui_MainRaporKontrol()
        self.ui.setupUi(self)

        self.ignored_exceptions=(NoSuchElementException,StaleElementReferenceException,)
        self.home = str(Path.home())+'\\rapor_kontrol'

        #kolonlar;
        self.lokasyon = DataFrame()
        self.kullaniciAdi = DataFrame()
        self.isyeriKodu = DataFrame()
        self.sifre = DataFrame()

        self.GirisBaslikRESelf = ''  # hatalı giriş uyarı mesajı kontrol

        self.Vakalar = ['Is Kazasi', 'Hastalik', 'Analik']

        #tablo1 mimarisi;
        self.ColumnName = ['Lokasyon', 'TC Kimlik No', 'Ad Soyad', 'Vaka', 'Takip No', 'Sıra No', 'Başlangıç', 'Işbaşı/Kontrol', 'Ceza Durumu', 'Sicil No']
        self.ColumnWidth = [435, 140, 250, 110, 210, 70, 110, 110, 385, 385]
        stylesheet = "::section{Background-color:rgb(c);border-radius:16px;font: 75 10pt "'MS Shell Dlg 2'";}"
        self.ui.tableProducts.horizontalHeader().setStyleSheet(stylesheet)
        self.ui.tableProducts.setColumnCount(len(self.ColumnName))
        self.ui.tableProducts.setHorizontalHeaderLabels(self.ColumnName)
        for index, width in enumerate(self.ColumnWidth):
            self.ui.tableProducts.setColumnWidth(index,width)

        #tablo1 mimarisi
        self.ColumnName2 = ['Lokasyon', 'Rapor Sonucu', 'Sicil No']
        self.ColumnWidth2 = [435, 140, 385]
        self.ui.tableProducts_2.horizontalHeader().setStyleSheet(stylesheet)
        self.ui.tableProducts_2.setColumnCount(len(self.ColumnName2))
        self.ui.tableProducts_2.setHorizontalHeaderLabels(self.ColumnName2)
        for index2, width2 in enumerate(self.ColumnWidth2):
            self.ui.tableProducts_2.setColumnWidth(index2,width2)
        
        #signal slots
        self.ui.btnSelectFile.clicked.connect(self.DosyaSec)
        self.ui.btnStart.clicked.connect(self.Baslat)
        self.ui.actionExcel.triggered.connect(self.ExcelAktar)
        self.ui.actionSil.triggered.connect(self.TabloSil)

    def DosyaSec(self):
        self.ui.txtDosyaYolLoad.clear()
        dosya_yol, _  = QFileDialog.getOpenFileName(self, "Dosya Aç", "", "Excel (*.xls *.xlsx)")
        path = Path(dosya_yol)
        dosya=path.name
        dosyaYol = path.resolve()
        
        if dosya_yol:    
            try:
                df= DosyaIslem(dosyaYol)
                self.ui.txtDosyaYolLoad.insert(dosya)

                self.lokasyon = df.DosyaOnIslem()[3]
                self.kullaniciAdi = df.DosyaOnIslem()[0]
                self.isyeriKodu = df.DosyaOnIslem()[1]
                self.sifre = df.DosyaOnIslem()[2]

            except Exception as err:
                baslik = 'HATALI DOSYA...!'
                metin = 'Dosyada Hatalı Veriler Var..\n\nVerilerinizi Aşağıdaki Detaylara Göre Girmelisiniz!'
                metin2 = "Sadece aşağıdaki bilgileri giriniz: \nLokasyon\nKullanıcı Kodu\nİşyer Şifresi"
                self.MesajBox(baslik, metin, metin2)
                print(err)
                self.ui.txtDosyaYolLoad.clear()
    
    def ImageText(self): # tesseractExe
        fn = 'Key.png'
        path = self.home+r'\{}'.format(fn)
        pytesseract.tesseract_cmd = self.home+'\\tesseractExe\\tesseract.exe'
        image = imread(path)
        data = pytesseract.image_to_string(image, lang='eng', config='--psm 6')

        return data
    
    def Baslat(self):
        baslik = 'İşlem Başlatılsın mı?'
        metin = 'İşlem Başlasın mı?\nDikkat..!\nİşlem Bitmeden Uygulamayı Kapatmayınız'
        evet = 'Evet'
        hayir = 'Hayır'
        result = self.MesajBoxSoru(baslik, metin, evet, hayir)
        if result == evet:
            self.Giris()

    def Giris(self): # kazıma / tablolara aktarma
        
        # options = Options()
        # options.set_headless(headless=True)     
        # self.driver = webdriver.Chrome(r'C:\Users\umit\data science\PyQt5\_Rapor_Kontrol\chromedriver.exe', options=opt)
        dosya = self.ui.txtDosyaYolLoad.text()
        if dosya:
            self.site = 'https://uyg.sgk.gov.tr/vizite/welcome.do'

            try:
                self.showMinimized()
                opt = webdriver.ChromeOptions()
                opt.add_experimental_option('excludeSwitches', ['enable-logging'])
                opt.add_argument("--start-maximized")

                #chromedriver_autoinstaller.install()
                self.driver = webdriver.Chrome(self.home+r'\chromedriver.exe', options=opt)
                self.driver.get(self.site)
                sleep(2)
                kullaniciADIxpath = '/html/body/table[2]/tbody/tr/td/form/table/tbody/tr[2]/td[3]/input'
                IsyeriKODUzpath = '/html/body/table[2]/tbody/tr/td/form/table/tbody/tr[2]/td[5]/b/input'
                kullaniciSIFRExpath = '/html/body/table[2]/tbody/tr/td/form/table/tbody/tr[3]/td[3]/input'
                guvenlikAnahtarixpath = '/html/body/table[2]/tbody/tr/td/form/table/tbody/tr[4]/td[3]/input'

                xPathKey = '/html/body/table[2]/tbody/tr/td/form/table/tbody/tr[5]/td[3]/img'
                           
                for ad, kod, sifre, lok in zip(self.kullaniciAdi, self.isyeriKodu, self.sifre, self.lokasyon):
                    try:
                        number, decimal = sifre.split(".")
                        if decimal == "0":
                            sifre = int(number)
                    except:
                        sifre = str(sifre)
                    
                    with open(self.home+'\Key'+'.png', 'wb') as file:
                        file.write(self.driver.find_element(By.XPATH, xPathKey).screenshot_as_png)
                        
                    key = self.ImageText()
                    kullaniciADIinput = self.driver.find_element(By.XPATH,kullaniciADIxpath)
                    IsyeriKODUinput = self.driver.find_element(By.XPATH, IsyeriKODUzpath)
                    kullaniciSIFREinput = self.driver.find_element(By.XPATH, kullaniciSIFRExpath)
                    guvenlikAnahtari = self.driver.find_element(By.XPATH, guvenlikAnahtarixpath)

                    while True:
                        try:
                            kullaniciADIinput.send_keys(str(ad))
                            sleep(0.1)
                            IsyeriKODUinput.send_keys(str(kod))
                            sleep(0.1)
                            kullaniciSIFREinput.send_keys(str(sifre))
                            sleep(0.1)
                            guvenlikAnahtari.send_keys(key)
                            sleep(1.1)
                        except:
                            continue
                
                        finally:
                            rapor_soup=BeautifulSoup(self.driver.page_source,"html.parser")
                            HataliGiris = [element.text for element in rapor_soup.find_all("td", "message")]
                            GirisBaslik = [element.text for element in rapor_soup.find_all("tr", "headerRow")]
                            GirisBaslikRE = re.search('([A-Z])\w+', str(GirisBaslik))
                            self.GirisBaslikRESelf = GirisBaslikRE.group()

                            if HataliGiris !=[] : #eğer giriş anahtarı hatalıysa
                                while HataliGiris != [] and self.GirisBaslikRESelf=='Kullanıcı': #hatalı giriş olduğu sürece
                                    with open(self.home+'\Key'+'.png', 'wb') as file:
                                        file.write(self.driver.find_element(By.XPATH, xPathKey).screenshot_as_png)
                                    keyT = self.ImageText()
                                    
                                    kullaniciADIinput = self.driver.find_element(By.XPATH,kullaniciADIxpath)
                                    IsyeriKODUinput = self.driver.find_element(By.XPATH, IsyeriKODUzpath)
                                    kullaniciSIFREinput = self.driver.find_element(By.XPATH, kullaniciSIFRExpath)
                                    guvenlikAnahtari = self.driver.find_element(By.XPATH, guvenlikAnahtarixpath)

                                    kullaniciADIinput.send_keys(str(ad))
                                    sleep(0.1)
                                    IsyeriKODUinput.send_keys(str(kod))
                                    sleep(0.1)
                                    kullaniciSIFREinput.send_keys(str(sifre))
                                    sleep(0.1)
                                    guvenlikAnahtari.send_keys(keyT)
                                    sleep(1.1)

                                    rapor_soup2=BeautifulSoup(self.driver.page_source,"html.parser")
                                    GirisBaslik2 = [element.text for element in rapor_soup2.find_all("tr", "headerRow")]
                                    GirisBaslikRE2 = re.search('([A-Z])\w+', str(GirisBaslik2))
                                    self.GirisBaslikRESelf = GirisBaslikRE2.group()
                            
                                break
                            
                        break

                    sicilNoXpath = '/html/body/table[2]/tbody/tr/td[1]/table[1]/tbody/tr/td/p'
                    sicilNoX = self.driver.find_element(By.XPATH, sicilNoXpath)
                    xz = sicilNoX.text
                    sc = xz[12:46].replace('-', ' ')

                    for vk in self.Vakalar:
                        self.IkinciAsamaTariheGoreRaporArama(lok, vk, sc)
                    
                    self.UcuncuAsamaArsiveGoreRaporArama(lok, sc)
                    
                    cikispath = '/html/body/table[2]/tbody/tr/td[1]/table[3]/tbody/tr[2]/td/table/tbody/tr/td[2]/a'
                    cikis = self.driver.find_element(By.XPATH, cikispath)
                    cikis.click()
                    sleep(0.5)
                self.driver.close()
                self.lokasyon = DataFrame()
                self.kullaniciAdi = DataFrame()
                self.isyeriKodu = DataFrame()
                self.sifre = DataFrame()
                self.ui.txtDosyaYolLoad.clear()
            
                
            except Exception as err :
                baslik = 'İNTERNET BAĞLANTISI YOK...!'
                metin = 'İnternet Bağlantınızı Kontrol Ediniz..\n'
                metin2 = str(err)
                self.MesajBox(baslik, metin, metin2)
                self.driver.close()
                print(err)
                return

    def IkinciAsamaTariheGoreRaporArama(self, lok, vaka, sc):
        tariheGoreRapor = self.driver.find_element(By.XPATH, '/html/body/table[2]/tbody/tr/td[1]/table[5]/tbody/tr[3]/td/table/tbody/tr/td[2]/a')
        tariheGoreRapor.click()

        #sleep(0.1)
        select = Select(self.driver.find_element(By.ID, 'vaka'))
        select.select_by_visible_text(vaka)

        #sleep(0.1)
        buGun = strftime('%x')
        TarihXpath = '/html/body/table[2]/tbody/tr/td[2]/form/table/tbody/tr[2]/td[2]/input'
        TarihXpathInput = self.driver.find_element(By.XPATH, TarihXpath)

        #sleep(0.1) 'WebDriver' object has no attribute 'find_element_by_id'
        TarihXpathInput.send_keys(buGun)

        #sleep(0.1)
        RaporAraXpath = '/html/body/table[2]/tbody/tr/td[2]/form/table/tbody/tr[4]/td/input'
        RaporAraXpathInput = self.driver.find_element(By.XPATH, RaporAraXpath)
        RaporAraXpathInput.click()

        #sleep(0.5)
        rapor_soup=BeautifulSoup(self.driver.page_source,"html.parser")
        KayitYokMesaj = [element.text for element in rapor_soup.find_all("td", "message")]
        
        #sleep(0.1)
        Mesaj = [[]]
        MesajYok = []
        
        if KayitYokMesaj !=[]:
            MesajYok.append(lok)
            MesajYok.append(vaka+' YOK')
            MesajYok.append(sc)
            rowCount2 = self.ui.tableProducts_2.rowCount()
            self.ui.tableProducts_2.insertRow(rowCount2)

            for w, z in enumerate(MesajYok):
                if z!="":
                    self.ui.tableProducts_2.setItem(rowCount2, w, QTableWidgetItem(z))

        else:
            KayitlarMesaj = [element.text for element in rapor_soup.find_all("td", "labelsmall9")]

            parcala = [KayitlarMesaj[x:x+9] for x in range(0, len(KayitlarMesaj), 9)]
            parcala = parcala[1:]

            for li in parcala:
                li = li[1:]
                listt = []
                listt.append(lok)
                for j in li:
                    strp = " ".join(j.split())
                    listt.append(strp)
                listt.append(sc)
                Mesaj.append(listt)
                
            Mesaj = Mesaj[1:]

            for li in Mesaj:
                rowCount = self.ui.tableProducts.rowCount()
                self.ui.tableProducts.insertRow(rowCount)
                for x, y in enumerate(li):
                    t = re.match('(\d{4})(-)(\d{2})(-)(\d{2})', y)
                    if t:
                        t = t.group()
                        if y==t:
                            y = y[8:10]+'.'+y[5:7]+'.'+y[0:4]
                    if y!="":
                        self.ui.tableProducts.setItem(rowCount, x, QTableWidgetItem(y))
         
    def UcuncuAsamaArsiveGoreRaporArama(self, lok, sc):
        arsiveGoreRapor = self.driver.find_element(By.XPATH, '/html/body/table[2]/tbody/tr/td[1]/table[5]/tbody/tr[6]/td/table/tbody/tr/td[2]/a')
        arsiveGoreRapor.click()

        #sleep(0.1)
        buGun = strftime('%x')
        TarihBaslangicXpath = '/html/body/table[2]/tbody/tr/td[2]/form/table/tbody/tr[2]/td[1]/input'
        TarihBaslangicXpathInput = self.driver.find_element(By.XPATH, TarihBaslangicXpath)

        #sleep(0.1)
        baslangicT = '01.01.2000'
        TarihBaslangicXpathInput.send_keys(baslangicT)

        #sleep(0.1)
        buGun = strftime('%x')
        TarihBitisXpath = '/html/body/table[2]/tbody/tr/td[2]/form/table/tbody/tr[2]/td[2]/input'
        TarihBitisXpathInput = self.driver.find_element(By.XPATH, TarihBitisXpath)

        #sleep(0.1)
        TarihBitisXpathInput.send_keys(buGun)

        #sleep(0.1)
        RaporAraXpath = '/html/body/table[2]/tbody/tr/td[2]/form/table/tbody/tr[3]/td/input'
        RaporAraXpathInput = self.driver.find_element(By.XPATH, RaporAraXpath)
        RaporAraXpathInput.click()

        #sleep(0.5)
        rapor_soup=BeautifulSoup(self.driver.page_source,"html.parser")
        KayitYokMesaj = [element.text for element in rapor_soup.find_all("td", "message")]
        
        #sleep(0.1)
        Mesaj = [[]]
        MesajYok = []
        
        if KayitYokMesaj !=[]:
            MesajYok.append(lok)
            MesajYok.append('Arşiv'+' YOK')
            MesajYok.append(sc)

            rowCount2 = self.ui.tableProducts_2.rowCount()
            self.ui.tableProducts_2.insertRow(rowCount2)

            for w, z in enumerate(MesajYok):
                if z!="":
                    self.ui.tableProducts_2.setItem(rowCount2, w, QTableWidgetItem(z))

        else:
            KayitlarMesaj = [element.text for element in rapor_soup.find_all("td", "labelsmall9")]
            
            parcala = [KayitlarMesaj[x:x+6] for x in range(0, len(KayitlarMesaj), 6)]
            parcala = parcala[1:]

            for li in parcala:
                li = li[1:]
                listt = []
                listt.append(lok)
                for j in li:
                    strp = " ".join(j.split())
                    listt.append(strp)
                listt.append(sc)
                Mesaj.append(listt)
                
            Mesaj = Mesaj[1:]

            ArsivMessage = []
            for c in Mesaj:
                c[3:3]='-'*3
                ArsivMessage.append(c)
            
            for li in ArsivMessage:
                rowCount = self.ui.tableProducts.rowCount()
                self.ui.tableProducts.insertRow(rowCount)
                for x, y in enumerate(li):
                    t = re.match('(\d{4})(-)(\d{2})(-)(\d{2})', y)
                    if t:
                        t = t.group()
                        if y==t:
                            y = y[8:10]+'.'+y[5:7]+'.'+y[0:4]
                    if y!="":
                        self.ui.tableProducts.setItem(rowCount, x, QTableWidgetItem(y))
    
    def TablodakiVeriler(self): #sadece ilk tablo için
        basliklar = []
        
        for i in range(self.ui.tableProducts.model().columnCount()):
            basliklar.append(self.ui.tableProducts.horizontalHeaderItem(i).text())
        
        df = DataFrame(columns=basliklar)

        for j in range(self.ui.tableProducts.rowCount()):
            for clm in range(self.ui.tableProducts.columnCount()):
                obj = self.ui.tableProducts.item(j,clm)
                if obj is not None and obj.text() != '':
                    df.at[j, basliklar[clm]] = self.ui.tableProducts.item(j, clm).text()
        return df

    def ExcelAktar(self): #sadece ilk tablo için
        rowCount = self.ui.tableProducts.rowCount()
        if rowCount > 0:
            baslik = 'Excel Dosyası'
            metin = 'Tablodaki Veriler Excele aktarılsın mı?'
            evet = 'Evet'
            hayir = 'Hayır'
            result = self.MesajBoxSoru(baslik, metin, evet, hayir)
            if result == evet:
                df = self.TablodakiVeriler()
                Xcpath = path.join(path.expanduser("~"), "Desktop", "Vizite_Rapor.xlsx")
                df.to_excel(Xcpath, index=False)
                baslik = 'EXCEL DOSYASI'
                metin = 'Vizite_Rapor.xlsx Masaüstünde Oluşturuldu'
                ok = 'Ok.'
                self.MesajBoxWarning(baslik, metin, ok)
        
        else:
            return

    def TabloSil(self): #iki tablo da
        rowCount = self.ui.tableProducts.rowCount()
        rowCount2 = self.ui.tableProducts_2.rowCount()
        if rowCount or rowCount2 > 0 :
            baslik = 'TABLO SİLME İŞLEMİ!!'
            metin = 'Tablodaki Veriler Silinsin mi?'
            evet = 'Evet'
            hayir = 'Hayır'
            result = self.MesajBoxSoru(baslik, metin, evet, hayir)
            if result == evet:
                while (self.ui.tableProducts.rowCount() > 0):
                    self.ui.tableProducts.removeRow(0)

                    self.ui.tableProducts.setColumnCount(len(self.ColumnName))
                    self.ui.tableProducts.setHorizontalHeaderLabels(self.ColumnName)
                    for index, width in enumerate(self.ColumnWidth):
                        self.ui.tableProducts.setColumnWidth(index,width)

                while (self.ui.tableProducts_2.rowCount() > 0):
                    self.ui.tableProducts_2.removeRow(0)

                    self.ui.tableProducts_2.setColumnCount(len(self.ColumnName2))
                    self.ui.tableProducts_2.setHorizontalHeaderLabels(self.ColumnName2)
                    for index2, width2 in enumerate(self.ColumnWidth2):
                        self.ui.tableProducts.setColumnWidth(index2,width2)
                    
        elif rowCount == 0:
            return

    def MesajBox(self, baslik, metin, metin2):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Question)
        msg.setWindowTitle(baslik)
        msg.setBaseSize
        
        mesajMetin  = '<pre style="font-size:12pt; color: #01040a;">{}<figure>'.format(metin)
        
        
        msg.setText(mesajMetin)
        msg.setStandardButtons(QMessageBox.Ok)
        msg.setDetailedText(metin2)
        msg.exec_()
    
    def MesajBoxSoru(self, baslik, metin, evet, hayir):
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Question)
        mesajMetin  = '<pre style="font-size:12pt; color: #064e9a;">{}<figure>'.format(metin)
        msg.setWindowTitle(baslik)
        
        msg.setText(mesajMetin)
        msg.setStandardButtons(QMessageBox.Yes|QMessageBox.No)
        buttonY = msg.button(QMessageBox.Yes)
        buttonY.setText(evet)
        buttonN = msg.button(QMessageBox.No)
        buttonN.setText(hayir)
        msg.exec_()
        if msg.clickedButton() == buttonY:
            return evet
        
        if msg.clickedButton() == buttonN:
            return hayir

    def MesajBoxWarning(self, baslik, metin, ok):
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Warning)
        mesajMetin  = '<pre style="font-size:12pt; color: #01040a;">{}<figure>'.format(metin)
        msg.setWindowTitle(baslik)

        msg.setText(mesajMetin)
        msg.setStandardButtons(QMessageBox.Ok)
        buttonOk = msg.button(QMessageBox.Ok)
        buttonOk.setText(ok)
        msg.exec_()

    def closeEvent(self, aksiyon):
        baslik = 'ÇIKIŞ'
        metin = "UYGULAMA KAPATILSIN MI?"
        evet = 'Evet'
        hayir = 'Hayır'
        result = self.MesajBoxSoru(baslik, metin, evet, hayir)
        if result == evet:
            #QMainWindow.closeEvent(self, event)
            while (self.ui.tableProducts.rowCount() > 0):
                    self.ui.tableProducts.removeRow(0)
            aksiyon.accept()
            #self.close()
        else:
            aksiyon.ignore()
    
    def keyPressEvent(self, Escape):
        if Escape.key() == Qt.Key_Escape:
            self.close()
        else:
            super(RaporKontrol, self).keyPressEvent(Escape)


def app():
    app = QApplication(argv)
    app.setStyle('Fusion')
    win = RaporKontrol()
    win.show()
    exit(app.exec())


app()