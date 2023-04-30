# -*- coding: utf-8 -*-

# Veri ön işleme

from pandas import read_excel, NaT


class DosyaIslem:

    def __init__(self, Dosya):
        self.Dosya = Dosya

    def DosyaOnIslem(self):

        df= read_excel(self.Dosya, header=None, engine='openpyxl')
        df = df.apply(lambda x: x.replace(r'^\s*$', NaT, regex=True) if x.dtype == "object" else x)
        df = df.apply(lambda x: x.replace(nan, NaT, regex=True) if x.dtype == "float64" else x)
        df = df.apply(lambda x: x.replace(nan, NaT, regex=True) if x.dtype == "float32" else x)
        df = df.apply(lambda x: x.replace(nan, NaT, regex=True) if x.dtype == "int" else x)
        df.dropna(how='all',axis=1, inplace=True)
        df.dropna(how='all',axis=0, inplace=True)
        df=df.rename(index={j: i for i, j in enumerate(df.index)})
        df = df.dropna(axis=1)
        Index = len(df.index)
        baslikkontrol = [p for p in df.values[0] if isinstance(p, float) or isinstance(p, int)] #başlıkları varsa boş liste

        if baslikkontrol==[]: #başlık var ise
            df = df.iloc[1:Index]
            df=df.rename(index={j: i for i, j in enumerate(df.index)})

            df.columns = [0,1,2,3]
            lokasyon = df[0].values
            kullaniciKodu = df[1].apply(int).values
            isyeriKodu = df[2].apply(int).values
            sifre = df[3].astype(str, errors = 'ignore')

            return kullaniciKodu, isyeriKodu, sifre, lokasyon, df
        
        else:
            df.columns = [0,1,2,3]
            lokasyon = df[0].values
            kullaniciKodu = df[1].apply(int).values
            isyeriKodu = df[2].apply(int).values
            sifre = df[3].astype(str, errors = 'ignore')
            
            return kullaniciKodu, isyeriKodu, sifre, lokasyon, df

# dosya = "sifreler.xlsx"

# file = DosyaIslem(dosya)

# print(file.DosyaOnIslem()[0])
