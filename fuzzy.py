# Property of Iyank Do Not Steal or Make a copy of it.
# Built using PyCharm

import pandas as PD
import random as RD

#fungsi u
def fungsiu(x,a,b):
    return ((x-a)/(b-a))

#read excel convert from dataframe to list
def dataread():
    data = PD.read_excel('Mahasiswa.xls')
    temp = PD.DataFrame(data, columns=['Id','Penghasilan','Pengeluaran'])
    list_data = temp.values.tolist()
    return list_data

#penghitungan nilai fuzy pengeluaran
def fuzyluar(luar):
    temp = []
    if(luar < 2):
        temp.append(["Rendah",1])
        temp.append(["Menengah",0])
    elif(luar == 5):
        temp.append(["Menengah",1])
        temp.append(["Tinggi",0])
    elif(luar >= 8):
        temp.append(["Tinggi",1])
        temp.append(["Menengah",0])
    elif(2 <= luar < 5):
        hasil = fungsiu(luar,2,5)
        temp.append(["Menengah",hasil])
        temp.append(["Rendah",1-hasil])
    elif(5 <= luar < 8 ):
        hasil = fungsiu(luar,5,8)
        temp.append(["Tinggi",hasil])
        temp.append(["Menengah",1-hasil])

    return temp

#penghitungan nilai fuzy penghasilan
def fuzyhasil(hasilan):
    temp = []
    if (hasilan < 5):
        temp.append(["Rendah", 1])
        temp.append(["Menengah", 0])
    elif(7 <= hasilan <= 12):
        temp.append(["Menengah", 1])
        temp.append(["Tinggi", 0])
    elif(hasilan >= 15):
        temp.append(["Tinggi", 1])
        temp.append(["Menengah",0])
    elif(5 <= hasilan < 7):
        hasil = fungsiu(hasilan, 5, 7)
        temp.append(["Menengah", hasil])
        temp.append(["Rendah", 1 - hasil])
    elif(12 < hasilan < 15):
        hasil = fungsiu(hasilan, 12, 15)
        temp.append(["Tinggi", hasil])
        temp.append(["Menengah", 1 - hasil])

    return temp

def inferensi(fp,fl):
    nilai_tinggi = []
    nilai_rendah = []
    nk = []

    if((fp[0][0] == "Rendah") and (fl[0][0]) == "Rendah"):
        mini = min(fp[0][1],fl[0][1])
        nilai_rendah.append(mini)
    if((fp[0][0] == "Rendah") and (fl[0][0]) == "Menengah"):
        mini = min(fp[0][1], fl[0][1])
        nilai_tinggi.append(mini)
    if((fp[0][0] == "Rendah") and (fl[0][0]) == "Tinggi"):
        mini = min(fp[0][1], fl[0][1])
        nilai_tinggi.append(mini)
    if((fp[0][0] == "Menengah") and (fl[0][0]) == "Rendah"):
        mini = min(fp[0][1], fl[0][1])
        nilai_rendah.append(mini)
    if((fp[0][0] == "Menengah") and (fl[0][0]) == "Menengah"):
        mini = min(fp[0][1], fl[0][1])
        nilai_rendah.append(mini)
    if((fp[0][0] == "Menengah") and (fl[0][0]) == "Tinggi"):
        mini = min(fp[0][1], fl[0][1])
        nilai_tinggi.append(mini)
    if((fp[0][0] == "Tinggi") and (fl[0][0]) == "Rendah"):
        mini = min(fp[0][1], fl[0][1])
        nilai_rendah.append(mini)
    if((fp[0][0] == "Tinggi") and (fl[0][0]) == "Menengah"):
        mini = min(fp[0][1], fl[0][1])
        nilai_rendah.append(mini)
    if((fp[0][0] == "Tinggi") and (fl[0][0]) == "Tinggi"):
        mini = min(fp[0][1], fl[0][1])
        nilai_rendah.append(mini)
    if((fp[0][0] == "Rendah") and (fl[1][0]) == "Rendah"):
        mini = min(fp[0][1], fl[1][1])
        nilai_rendah.append(mini)
    if((fp[0][0] == "Rendah") and (fl[1][0]) == "Menengah"):
        mini = min(fp[0][1], fl[1][1])
        nilai_tinggi.append(mini)
    if((fp[0][0] == "Rendah") and (fl[1][0]) == "Tinggi"):
        mini = min(fp[0][1], fl[1][1])
        nilai_tinggi.append(mini)
    if((fp[0][0] == "Menengah") and (fl[1][0]) == "Rendah"):
        mini = min(fp[0][1], fl[1][1])
        nilai_rendah.append(mini)
    if((fp[0][0] == "Menengah") and (fl[1][0]) == "Menengah"):
        mini = min(fp[0][1], fl[1][1])
        nilai_rendah.append(mini)
    if((fp[0][0] == "Menengah") and (fl[1][0]) == "Tinggi"):
        mini = min(fp[0][1], fl[1][1])
        nilai_tinggi.append(mini)
    if((fp[0][0] == "Tinggi") and (fl[1][0]) == "Rendah"):
        mini = min(fp[0][1], fl[1][1])
        nilai_rendah.append(mini)
    if((fp[0][0] == "Tinggi") and (fl[1][0]) == "Menengah"):
        mini = min(fp[0][1], fl[1][1])
        nilai_rendah.append(mini)
    if((fp[0][0] == "Tinggi") and (fl[1][0]) == "Tinggi"):
        mini = min(fp[0][1], fl[1][1])
        nilai_rendah.append(mini)
    if((fp[1][0] == "Rendah") and (fl[0][0]) == "Rendah"):
        mini = min(fp[1][1], fl[0][1])
        nilai_rendah.append(mini)
    if((fp[1][0] == "Rendah") and (fl[0][0]) == "Menengah"):
        mini = min(fp[1][1], fl[0][1])
        nilai_tinggi.append(mini)
    if((fp[1][0] == "Rendah") and (fl[0][0]) == "Tinggi"):
        mini = min(fp[1][1], fl[0][1])
        nilai_tinggi.append(mini)
    if((fp[1][0] == "Menengah") and (fl[0][0]) == "Rendah"):
        mini = min(fp[1][1], fl[0][1])
        nilai_rendah.append(mini)
    if((fp[1][0] == "Menengah") and (fl[0][0]) == "Menengah"):
        mini = min(fp[1][1], fl[0][1])
        nilai_rendah.append(mini)
    if((fp[1][0] == "Menengah") and (fl[0][0]) == "Tinggi"):
        mini = min(fp[1][1], fl[0][1])
        nilai_tinggi.append(mini)
    if((fp[1][0] == "Tinggi") and (fl[0][0]) == "Rendah"):
        mini = min(fp[1][1], fl[0][1])
        nilai_rendah.append(mini)
    if((fp[1][0] == "Tinggi") and (fl[0][0]) == "Menengah"):
        mini = min(fp[1][1], fl[0][1])
        nilai_rendah.append(mini)
    if((fp[1][0] == "Tinggi") and (fl[0][0]) == "Tinggi"):
        mini = min(fp[1][1], fl[0][1])
        nilai_rendah.append(mini)
    if((fp[1][0] == "Rendah") and (fl[1][0]) == "Rendah"):
        mini = min(fp[1][1], fl[1][1])
        nilai_rendah.append(mini)
    if((fp[1][0] == "Rendah") and (fl[1][0]) == "Menengah"):
        mini = min(fp[1][1], fl[1][1])
        nilai_tinggi.append(mini)
    if((fp[1][0] == "Rendah") and (fl[1][0]) == "Tinggi"):
        mini = min(fp[1][1], fl[1][1])
        nilai_tinggi.append(mini)
    if((fp[1][0] == "Menengah") and (fl[1][0]) == "Rendah"):
        mini = min(fp[1][1], fl[1][1])
        nilai_rendah.append(mini)
    if((fp[1][0] == "Menengah") and (fl[1][0]) == "Menengah"):
        mini = min(fp[1][1], fl[1][1])
        nilai_rendah.append(mini)
    if((fp[1][0] == "Menengah") and (fl[1][0]) == "Tinggi"):
        mini = min(fp[1][1], fl[1][1])
        nilai_tinggi.append(mini)
    if((fp[1][0] == "Tinggi") and (fl[1][0]) == "Rendah"):
        mini = min(fp[1][1], fl[1][1])
        nilai_rendah.append(mini)
    if((fp[1][0] == "Tinggi") and (fl[1][0]) == "Menengah"):
        mini = min(fp[1][1], fl[1][1])
        nilai_rendah.append(mini)
    if((fp[1][0] == "Tinggi") and (fl[1][0]) == "Tinggi"):
        mini = min(fp[1][1], fl[1][1])
        nilai_rendah.append(mini)


    if (len(nilai_rendah) == 0):
        nk.append(0)
    else:
        nk.append(max(nilai_rendah))

    if(len(nilai_tinggi) == 0):
        nk.append(0)
    else:
        nk.append(max(nilai_tinggi))

    return nk

def defuzzy(nk1,nk2,randlist):
    rendah  = []
    tinggi  = []
    miring  = []
    sum_rendah = 0
    sum_tengah = 0
    sum_tinggi = 0
    temp1      = 0

    for i in randlist:
        if (i <= 40):
            rendah.append(i)
        elif (40 < i <= 60):
            miring.append(i)
        elif (i > 60):
            tinggi.append(i)

    for d in rendah:
        sum_rendah += d

    for i in range(len(miring)):
        if (nk1 <= nk2):
            sum_tengah += miring[i] * (miring[i]/100)
            temp1 += (miring[i]/100)
        else:
            sum_tengah += -((miring[i]/100)-1) * miring[i]
            temp1 += -((miring[i]/100)-1) * miring[i]
    for d in tinggi:
        sum_tinggi += d

    sum_rendah = sum_rendah*nk1
    sum_tinggi = sum_tinggi*nk2
    hasil = (sum_rendah + sum_rendah + sum_tinggi)
    tempp= (nk1 * len(rendah)) + temp1 + (nk2 * len(tinggi))

    return hasil/tempp

#main
#inisialiasi dictionary/variable

data = dataread()
data_mhs_pengeluaran  = []
data_mhs_penghasilan  = []
data_inferensi        = []
rand_list = RD.sample(range(0,100),10)
rand_list.sort()
templ = []
tempp = []
nk    = []
nKa   = []
list_akhir = []
list_to_convert = []

for d in data:
    templ = fuzyluar(d[2])
    tempp = fuzyhasil(d[1])
    templ.append(d[0])
    tempp.append(d[0])
    data_mhs_pengeluaran.append(templ)
    data_mhs_penghasilan.append(tempp)

for i in range(100):
    temp = inferensi(data_mhs_penghasilan[i], data_mhs_pengeluaran[i])
    nk.append(temp)

for i in range(100):
    temp = defuzzy(nk[i][0],nk[i][1],rand_list)
    nKa.append([i+1,temp])

list_akhir = sorted(nKa,key=lambda x: -x[1])

count = 0
while (count < 20):
    list_to_convert.append(list_akhir[count])
    count += 1

output_col = [
    "ID",
    "Nilai_Kelayakan"
]

file = PD.DataFrame(columns = output_col,data = list_to_convert)
file.to_excel('Bantuan.xls', engine='xlwt')
