import pandas as pd

def luaran():
    #akses file masukan.xlsx
    data = pd.read_excel('input.xlsx')
    
    #definisi variabel
    id = data["ID"]
    minat = data["Minat"]
    bakat = data["Bakat"]

    hasil = []

    for x in range(100):
        arr_minat = [0,0,0,0,0]
        arr_bakat = [0,0,0]

        rendah = kurang = normal = tinggi = sangatTinggi = 0
        pemula = standar = ahli = 0

        #fuzzification
        if minat[x] <= 10:
            rendah = 1
            arr_minat[0] = rendah
        elif minat[x] > 10 and minat[x] < 30:
            rendah = (30 - minat[x])/20
            kurang = (minat[x] - 10)/20
            arr_minat[0] = rendah
            arr_minat[1] = kurang
        elif minat[x] >= 30 and minat[x] <= 40:
            kurang = 1
            arr_minat[1] = kurang
        elif minat[x] > 40 and minat[x] < 60:
            kurang = (minat[x] - 40)/20
            normal = (60 - minat[x])/20
            arr_minat[1] = kurang
            arr_minat[2] = normal
        elif minat[x] >= 60 and minat[x] <=70:
            normal = 1
            arr_minat[2] = normal
        elif minat[x] > 70 and minat[x] < 80:
            normal = (80 - minat[x])/10
            tinggi = (minat[x] - 70)/10
            arr_minat[2] = normal
            arr_minat[3] = tinggi
        elif minat[x] >= 80 and minat[x] <= 85:
            tinggi = 1
            arr_minat[3] = tinggi
        elif minat[x] > 85 and minat[x] < 90:
            tinggi = (90 - minat[x])/5
            sangatTinggi = (minat[x] - 85)/5
            arr_minat[3] = tinggi
            arr_minat[4] = sangatTinggi
        elif minat[x] >= 90:
            sangatTinggi = 1
            arr_minat[4] = sangatTinggi

        if bakat[x] <= 10:
            pemula = 1
            arr_bakat[0] = pemula
        elif bakat[x] > 10 and bakat[x] < 20:
            pemula = (20 - bakat[x])/10
            standar = (bakat[x] - 10)/10
            arr_bakat[0] = pemula
            arr_bakat[1] = standar
        elif bakat[x] >= 20 and bakat[x] <= 40:
            standar = 1
            arr_bakat[1] = standar
        elif bakat[x] > 40 and bakat[x] < 50:
            standar = (50 - bakat[x])/10
            ahli = (bakat[x] - 40)/10
            arr_bakat[1] = standar
            arr_bakat[2] = ahli
        elif bakat[x] >= 50:
            ahli = 1
            arr_bakat[2] = ahli
        
        #inference
        L = []
        if arr_bakat[0] == pemula and arr_minat[0] == rendah:
            L.append(min(arr_bakat[0],arr_minat[0]))
        if arr_bakat[0] == pemula and arr_minat[1] == kurang:
            L.append(min(arr_bakat[0],arr_minat[1]))
        if arr_bakat[0] == pemula and arr_minat[2] == normal:
            L.append(min(arr_bakat[0],arr_minat[2]))
        if arr_bakat[0] == pemula and arr_minat[3] == tinggi:
            L.append(min(arr_bakat[0],arr_minat[3]))
        if arr_bakat[0] == pemula and arr_minat[4] == sangatTinggi:
            L.append(min(arr_bakat[0],arr_minat[4]))
        if arr_bakat[1] == standar and arr_minat[0] == rendah:
            L.append(min(arr_bakat[1],arr_minat[0]))
        max_L = max(L)

        M = []
        if arr_bakat[1] == standar and arr_minat[1] == kurang:
            M.append(min(arr_bakat[1],arr_minat[1]))
        if arr_bakat[1] == standar and arr_minat[2] == normal:
            M.append(min(arr_bakat[1],arr_minat[2]))
        if arr_bakat[1] == standar and arr_minat[3] == tinggi:
            M.append(min(arr_bakat[1],arr_minat[3]))
        if arr_bakat[1] == standar and arr_minat[4] == sangatTinggi:
            M.append(min(arr_bakat[1],arr_minat[4]))
        max_M = max(M)

        S = []
        if arr_bakat[2] == ahli and arr_minat[0] == rendah:
            S.append(min(arr_bakat[2],arr_minat[0]))
        if arr_bakat[2] == ahli and arr_minat[1] == kurang:
            S.append(min(arr_bakat[2],arr_minat[1]))
        if arr_bakat[2] == ahli and arr_minat[2] == normal:
            S.append(min(arr_bakat[2],arr_minat[2]))
        if arr_bakat[2] == ahli and arr_minat[3] == tinggi:
            S.append(min(arr_bakat[2],arr_minat[3]))
        if arr_bakat[2] == ahli and arr_minat[4] == sangatTinggi:
            S.append(min(arr_bakat[2],arr_minat[4]))
        max_S = max(S)

        #defuzzification
        bagi = max_L + max_M + max_S
        rs = ((max_S*20)+(max_M*40)+(max_L*60))/bagi
        hasil.append([id[x],rs])

        #output
        hasil_akhir = sorted(hasil, key = lambda x: x[1], reverse=True)
        hasil_xlsx = {'Kecocokan': hasil_akhir[:10]}

        result_xlsx = pd.DataFrame(hasil_xlsx, columns = ['Kecocokan'])
        result_xlsx.to_excel('luaran.xlsx')

#Program Utama
if __name__== "__main__":
    luaran()