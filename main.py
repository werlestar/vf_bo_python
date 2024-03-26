import csv
import xlrd
import openpyxl


# Hittar första strängen
def trunkera_transaktionsreferens(lista, kolumn):
    for i in range(1, len(lista)):
        if len(lista[i][kolumn]) > 12:
            for j in range(0, len(lista[i][kolumn])):
                if lista[i][kolumn][j].isdigit():
                    lista[i][kolumn] = lista[i][kolumn][j:]
                    break


# En funktion där två listors respektive kolumner jämförs för att hitta unika transaktioner
def hitta_unika(lista1, kolumn1, lista2, kolumn2):
    unika_vf = []
    unika_bo = []
    gemensamma_vf = []
    gemensamma_bo = []

    for x in range(1, len(lista1)):
        unik = True
        for y in range(1, len(lista2)):
            if lista1[x][kolumn1] == lista2[y][kolumn2]:
                if lista1[x][13] == lista2[y][12]:
                    unik = False
        if unik:
            unika_vf.append(lista1[x])
        else:
            if lista1[x][7] == '':
                continue
            gemensamma_vf.append(lista1[x])

    for x in range(1, len(lista2)):
        unik = True
        for y in range(1, len(lista1)):
            if lista2[x][kolumn2] == lista1[y][kolumn1]:
                if lista2[x][12] == lista1[y][13]:
                    unik = False
        if lista2[x][kolumn2] == '':
            unik = True
        if unik:
            unika_bo.append(lista2[x])
        else:
            gemensamma_bo.append(lista2[x])

    return unika_vf, unika_bo, gemensamma_vf, gemensamma_bo


# Ser till att alla siffror har samma antal decimaler och tar bort separationen för att omvandla till int
def städa_belopp(lista, kolumnlista):
    def körning(kolumn):
        for i in range(1, len(lista)):
            if '.' in lista[i][kolumn]:
                if lista[i][kolumn][-2] == '.':
                    lista[i][kolumn] += '0'
                lista[i][kolumn] = lista[i][kolumn].replace(',', '')
                lista[i][kolumn] = lista[i][kolumn].replace(' ', '')
                lista[i][kolumn] = lista[i][kolumn].replace('.', '')
            else:
                lista[i][kolumn] += '00'
    for x in range(0, len(kolumnlista)):
        körning(kolumnlista[x])


# Identifierar vilken kolumn strängen existerar vid
def hitta_kolumn(lista, sträng):
    for i in range(0, len(lista)):
        if lista[i] == sträng:
            return i


# Skapar en lista med kolumner som innehåller någon av rubrikerna i inputen
def skapa_hit_list(rad, rubriker):
    hit_list = []
    for i in range(0, len(rad)):
        if rad[i] in rubriker:
            hit_list.append(i)
    return hit_list


# Återinför separationen mellan decimalerna
def formatera_summa(n):
    if len(n) > 1:
        n = n[:-2] + '.' + n[-2:]
    return n


# Skapar en textrad för utskrift med rätt format
def skapa_rad(x, hit_list, lista, index):
    listan = [f'{lista[x][-1]};{lista[x][hit_list[index]][:-2]}.{lista[x][hit_list[0]][-2:]};']
    if index == 2:
        listan[0] += ';'

    for i in range(0, len(lista[x])):
        if i in hit_list:
            listan[0] += lista[x][i][:-2] + '.' + lista[x][i][-2:] + ';'
        else:
            listan[0] += lista[x][i] + ';'
    return listan


# Skapar och fyller det nya dokumentet med informationen från analsyen
def skriv_till_fil(vf, bo, vf_raden, bo_raden, gemensamma_vf, gemensamma_bo):
    with open('Resultat.csv', 'w', encoding='UTF8', newline='') as f:
        writer = csv.writer(f)
        summa = 0
        writer.writerow([f'VF ({len(vf)} av {len(vf)+len(gemensamma_vf)}) '])
        writer.writerow([f'Transaktions-ID;Värde; ;{vf_raden}'])
        for i in range(0, len(vf)):
            writer.writerow(skapa_rad(i, vf_hit_list, vf, 0))
            summa += int(vf[i][vf_hit_list[0]])

        writer.writerow([''])
        writer.writerow([f'Summa; {formatera_summa((str(summa)))}'])
        writer.writerow([''])

        summa = 0
        writer.writerow([f'BO ({len(bo)} av {len(bo)+len(gemensamma_bo)})'])
        writer.writerow([f'Transaktions-ID;Värde; ;{bo_raden}'])
        for i in range(0, len(bo)):
            writer.writerow(skapa_rad(i, bo_hit_list, bo, 2))
            summa += int(bo[i][bo_hit_list[2]])

        writer.writerow([''])
        writer.writerow([f'Summa; {formatera_summa((str(summa)))}'])
        writer.writerow([''])
        writer.writerow([''])
        writer.writerow([f'Gemensamma: {len(gemensamma_vf)+len(gemensamma_bo)}'])
        writer.writerow([''])
        writer.writerow([f'VF: ({len(gemensamma_vf)} av {len(vf)+len(gemensamma_vf)})'])
        for i in range(0, len(gemensamma_vf)):
            writer.writerow(skapa_rad(i, vf_hit_list, gemensamma_vf, 0))

        writer.writerow([''])
        writer.writerow([f'BO: ({len(gemensamma_bo)} av {len(bo)+len(gemensamma_bo)})'])
        for i in range(0, len(gemensamma_bo)):
            writer.writerow(skapa_rad(i, bo_hit_list, gemensamma_bo, 2))


# Läser in data från xls dokument
def vf_data_returner(string):
    data = []
    workbook = xlrd.open_workbook(string)
    worksheet = workbook.sheet_by_index(0)

    for row_idx in range(worksheet.nrows):
        row_data = []
        for col_idx in range(worksheet.ncols):
            cell_value = worksheet.cell_value(row_idx, col_idx)
            row_data.append(str(cell_value))
        data.append(row_data)

    return data


# Läser in data från xlsx fil
def bo_data_returner(string):
    xlsx_file = string
    workbook = openpyxl.load_workbook(xlsx_file)
    worksheet = workbook.active
    data = []

    for row in worksheet.iter_rows():
        row_data = []
        for cell in row:
            row_data.append(str(cell.value))
        data.append(row_data)

    return data


# Raderar rader som innehåller annat än den önskade informationen
def vf_trim(lista):
    for i in range(len(lista)-1, 0, -1):
        if lista[i][6] == '':
            if lista[i][7] == '':
                if lista[i][8] == '':
                    del lista[i]
    return lista


vf_data = vf_trim(vf_data_returner('VF.xls'))
bo_data = bo_data_returner('BO.xlsx')

städa_belopp(bo_data, [4, 5, 6])
städa_belopp(vf_data, [10, 11, 12])

vf_trans_kolumn = hitta_kolumn(vf_data[0], 'Filing code')
bo_trans_kolumn = hitta_kolumn(bo_data[0], 'Extern transaktionsreferens')

trunkera_transaktionsreferens(bo_data, bo_trans_kolumn)
vf_unika, bo_unika, gemens_vf, gemens_bo = hitta_unika(vf_data, vf_trans_kolumn, bo_data, bo_trans_kolumn)

vf_hit_list = skapa_hit_list(vf_data[0], ['Amount total', 'Purchase amount', 'Cashback amount'])
bo_hit_list = skapa_hit_list(bo_data[0], ['Belopp', 'Dricks', 'Totalt'])

vf_data[0] = ';'.join(vf_data[0])
vf_data[0] = vf_data[0][2:]
bo_data[0] = ';'.join(bo_data[0])
bo_data[0] = bo_data[0][1:]

skriv_till_fil(vf_unika, bo_unika, vf_data[0], bo_data[0], gemens_vf, gemens_bo)
# skapa_print(vf_unika, bo_unika, vf_data[0], bo_data[0], gemensamma)
