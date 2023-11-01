import pandas as pd
import os
import datetime as dt
import smtplib
from email.message import EmailMessage


def csv_name():
    folder = os.listdir("C:/Users/gokhan.kaya/OneDrive - Aster Textile/Desktop/BELGELERİM/PycharmProjects/PandasProjects/artesa_puantaj_kontrol")
    csv_str = []
    for search in folder:
        if "csv" in search:
            csv_str.append(search)
    return csv_str[0]


csv_data = pd.read_csv(csv_name(), sep=";")
df = pd.DataFrame(data=csv_data)
df = df.dropna(subset="AltFirma")



data_01 = list(df["mesaitarih"])
data_02 = list(df["Giriş"])
data_03 = list(df["Çıkış"])
data_04 = list(df["OFM"])
convert_01 = []
convert_02 = []
convert_03 = []
convert_04 = []


for con1, con2, con3, con4 in zip(data_01, data_02, data_03, data_04):
    convert_01.append(dt.datetime(int(con1[6:]), int(con1[3:5]), int(con1[:2])))

    if type(con2) == float:
        convert_02.append(None)
    else:
        convert_02.append(dt.time(int(con2[0:2]), int(con2[3:5]), int(con2[6:])))

    if type(con3) == float:
        convert_03.append(None)
    else:
        convert_03.append(dt.time(int(con3[0:2]), int(con3[3:5]), int(con3[6:])))

    if type(con4) == float:
        convert_04.append(None)
    else:
        convert_04.append(dt.time(int(con4[0:2]), int(con4[3:5])))

df["mesaitarih"] = convert_01
df["Giriş"] = convert_02
df["Çıkış"] = convert_03
df["OFM"] = convert_04


days = [aktar.weekday() for aktar in df["mesaitarih"]]
df["Günler"] = days


shift = df["Mesai Açıklama"]
status = []

content = {"Cerkezkoy 08:00-16:00": "Vardiya Mavi Yaka Hİ 1",
          "Cerkezkoy 16:00-24:00": "Vardiya Mavi Yaka Hİ 2",
          "Cerkezkoy 24:00-08:00": "Vardiya Mavi Yaka Hİ 3",
          "Cerkezkoy HT08": "Vardiya Mavi Yaka HS",
          "Cerkezkoy HT16": "Vardiya Mavi Yaka HS",
          "Cerkezkoy HT24": "Vardiya Mavi Yaka HS",
          "Cerkezkoy 08:00-17:15": "İdari Mavi Yaka Hİ",
          "Cerkezkoy Idari Ctesi Yeni": "İdari Mavi Yaka HS Cumartesi",
          "Cerkezkoy Idari Pazar": "İdari Mavi Yaka Hs Pazar",
          "Resmi Tatil Calısma tam": "Resmi Tatil",
          "Resmi Tatil Calısma yarım": "Resmi Tatil"}


for dd in shift:
   if dd in content.keys():
       status.append(content.get(dd))


df["Personel Durumu"] = status
df["Notlar"] = None

print(df.columns)

df = df[["sicilno", "Bölüm", "mesaitarih", "Giriş", "Çıkış", "OFM", "Mesai Açıklama", "İzin Açıklama", "Personel Durumu", "Notlar"]]




filter_01 = {1: (df["Giriş"].notna()) & (df["Çıkış"].isna()),
             2: (df["Giriş"].isna()) & (df["Çıkış"].notna()),
             3: (df["Personel Durumu"] == "İdari Mavi Yaka Hİ") & (df["Personel Durumu"] == "Vardiya Mavi Yaka Hİ 1") & (df["Personel Durumu"] == "Vardiya Mavi Yaka Hİ 2")
                & (df["Personel Durumu"] == "Vardiya Mavi Yaka Hİ 3") & (df["Giriş"].isna()) & (df["Çıkış"].isna()) & (df["İzin Açıklama"].isna()),
             4: (df["Personel Durumu"] == "İdari Mavi Yaka Hİ") & (df["Çıkış"] > dt.time(17, 45, 00)) & (df["OFM"] == dt.time(00, 00, 00)),
             5: (df["Personel Durumu"] == "Vardiya Mavi Yaka Hİ 1") & (df["Çıkış"] > dt.time(16, 30, 00)) & (df["OFM"] == dt.time(00, 00, 00)),
             6: (df["Personel Durumu"] == "Vardiya Mavi Yaka Hİ 2") & (df["Çıkış"] > dt.time(23, 59, 59)) & (df["OFM"] == dt.time(00, 00, 00)),
             7: (df["Personel Durumu"] == "Vardiya Mavi Yaka Hİ 3") & (df["Çıkış"] > dt.time(8, 30, 00)) & (df["OFM"] == dt.time(00, 00, 00)),
             8: (df["Personel Durumu"] == "İdari Mavi Yaka HS Cumartesi") & (df["Giriş"].notna()) & (df["Çıkış"] > dt.time(12, 45, 00)) & (df["OFM"] == dt.time(00, 00, 00)),
             9: (df["Personel Durumu"] == "İdari Mavi Yaka Hs Pazar") & (df["Giriş"].notna()) & (df["Çıkış"].notna()) & (df["OFM"] == dt.time(00, 00, 00)),
             10: (df["Personel Durumu"] == "Vardiya Mavi Yaka HS") & (df["Giriş"].notna()) & (df["Çıkış"].notna()) & (df["OFM"] == dt.time(00, 00, 00)),
             11: (df["Personel Durumu"] == "Resmi Tatil") & (df["Giriş"].notna()) & (df["Çıkış"].notna()) & (df["OFM"] == dt.time(00, 00, 00))}


frames = []
for ff in range(1, len(filter_01)+1):
    frames.append("frame" + str(ff))


tables = {}
for gg in range(len(frames)):
    tables[frames[gg]] = df



tables["frame1"] = df.loc[filter_01[1], :]
tables["frame1"].loc[:, "Notlar"] = "Personelin çıkış bilgisi eksik"
tablo_1 = pd.DataFrame(data=tables["frame1"])

tables["frame2"] = df.loc[filter_01[2], :]
tables["frame2"].loc[:, "Notlar"] = "Personelin giriş bilgisi eksik"
tablo_2 = pd.DataFrame(data=tables["frame2"])

tables["frame3"] = df.loc[filter_01[3], :]
tables["frame3"].loc[:, "Notlar"] = "Personelin İzin açıklaması eksik"
tablo_3 = pd.DataFrame(data=tables["frame3"])

tables["frame4"] = df.loc[filter_01[4], :]
tables["frame4"].loc[:, "Notlar"] = "Personelin haftaiçi OFM si eksik"
tablo_4 = pd.DataFrame(data=tables["frame4"])

tables["frame5"] = df.loc[filter_01[5], :]
tables["frame5"].loc[:, "Notlar"] = "Personelin haftaiçi OFM si eksik"
tablo_5 = pd.DataFrame(data=tables["frame5"])

tables["frame6"] = df.loc[filter_01[6], :]
tables["frame6"].loc[:, "Notlar"] = "Personelin haftaiçi OFM si eksik"
tablo_6 = pd.DataFrame(data=tables["frame6"])

tables["frame7"] = df.loc[filter_01[7], :]
tables["frame7"].loc[:, "Notlar"] = "Personelin haftaiçi OFM si eksik"
tablo_7 = pd.DataFrame(data=tables["frame7"])

tables["frame8"] = df.loc[filter_01[8], :]
tables["frame8"].loc[:, "Notlar"] = "Personelin haftasonu OFM si eksik"
tablo_8 = pd.DataFrame(data=tables["frame8"])

tables["frame9"] = df.loc[filter_01[9], :]
tables["frame9"].loc[:, "Notlar"] = "Personelin haftasonu OFM si eksik"
tablo_9 = pd.DataFrame(data=tables["frame9"])

tables["frame10"] = df.loc[filter_01[10], :]
tables["frame10"].loc[:, "Notlar"] = "Personelin haftasonu OFM si eksik"
tablo_10 = pd.DataFrame(data=tables["frame10"])

tables["frame11"] = df.loc[filter_01[11], :]
tables["frame11"].loc[:, "Notlar"] = "Personelin haftasonu OFM si eksik"
tablo_11 = pd.DataFrame(data=tables["frame11"])


pivot = pd.concat([tablo_1, tablo_2, tablo_3, tablo_4, tablo_5, tablo_6, tablo_7, tablo_8, tablo_9, tablo_10, tablo_11], axis=0)

pivot.to_excel("artesa_puantaj_kontrol.xlsx", sheet_name="Günlük Paporlar")


mail_subject = "PuantajKontrol"
mail_from = "pythonmailgonderim@gmail.com"
mail_to = ["yagmur.ulug@artesafabrics.com", "gokhan.kaya@astertextile.com", "ayten.guler@astertextile.com"]
mail_mesaj = "Merhaba, Lütfen ekteki puantaj özet dosyasını kendi tesisiniz bazında kontrol ediniz. İyi Çalışmalar."
appPassword = "clbj ioyn gmhy pqzp"

with open("artesa_puantaj_kontrol.xlsx", "rb") as f:
    file = f.read()
    file_name = f.name

mail = EmailMessage()
mail["Subject"] = mail_subject
mail["From"] = mail_from
mail["To"] = mail_to
mail.set_content(mail_mesaj)
mail.add_attachment(file, maintype="application", subtype="octet-stream", filename=file_name)

with smtplib.SMTP_SSL("smtp.gmail.com") as sent:
    sent.login("pythonmailgonderim@gmail.com", appPassword)
    sent.send_message(mail)
    sent.quit()
