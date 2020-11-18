import math
from openpyxl import load_workbook
import os
import pandas as pd
import re
import statistics

from Classes import *

cursosSel = []
sel = [["2", "3", "4", "5"], ["2", "3", "6", "7", "11"], ["2", "3", "4", "5", "6", "7", "8", "9"], ["2", "4", "6", "8", "10"]]
pesos = [0, 10, 9, 8, 6, 7, 5, 4, 3, 1, 2]
vagasPerc = [.4, .00966, .03234, .04784, .16016, .00966, .03234, .04784, .16016, .05, .05]


def listarCandidatos():
    arquivos = os.listdir(('C:\\temp'))
    cursos, candidatos, medias = [], [], []
    for arq in arquivos:
        nome = re.split(r'[-.]+', arq)
        cursosSel.append((nome[2], nome[3]))
        cursos.append(open('C:\\temp\\' + arq, encoding="utf8"))
    for cur in cursos:
        candidato = cur.read().split('\n')
        for linha in candidato:
            tx = linha.split(',')
            if (len(tx) > 1):
                candidatos.append(
                    Candidato(tx[5], tx[4], tx[2], tx[1], tx[3], tx[6], ajusteCotas(int(tx[6])), float(tx[12]), int(tx[13])))
    cur.close()
    return candidatos


def ajusteCotas(cotaInscricao):
    if cotaInscricao == 2:
        c = 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11
    elif cotaInscricao == 3:
        c = 1, 3, 5, 7, 9, 11
    elif cotaInscricao == 4:
        c = 1, 4, 5, 8, 9, 10
    elif cotaInscricao == 5:
        c = 1, 5, 9
    elif cotaInscricao == 6:
        c = 1, 6, 8, 9, 10, 11
    elif cotaInscricao == 7:
        c = 1, 7, 9, 11
    elif cotaInscricao == 8:
        c = 1, 8, 9, 10
    elif cotaInscricao == 9:
        c = 1, 9
    elif cotaInscricao == 10:
        c = 1, 10
    elif cotaInscricao == 11:
        c = 1, 11
    elif cotaInscricao == 12:
        c = 1, 2, 3, 4, 5, 6, 7, 8, 9, 10
    elif cotaInscricao == 13:
        c = 1, 3, 5, 7, 9
    elif cotaInscricao == 16:
        c = 1, 6, 8, 9, 10
    elif cotaInscricao == 17:
        c = 1, 7, 9
    else:
        c = 1,
    return c


def ordemChamada():
    ordem = []
    for num in range(11):
        ordem.append(Cota((num+1), vagasPerc[num], pesos[num]))
    return ordem


def destinarVagas(ordem, curso):
    ordem.sort(key=lambda cota: cota.peso, reverse=True)
    vagas = int(input(f'Qual é o total de vagas para a chamada do curso {curso[1]} - {curso[0]}? '))

    if vagas < 1:
        raise ValueError("Número de vagas inválidas")
    elif vagas == 1:
        ordem[10].vagas += 1
    else:
        vagasPublicas = math.ceil(vagas / 2)
        vagasUniversais = math.floor(vagas / 2)

        if vagasUniversais <= 3:
            ordem[10].vagas += 1
            vagasUniversais -= 1
            for i in range(8, 10, 1):
               if vagasUniversais > 0:
                    ordem[i].vagas += 1
                    vagasUniversais -= 1
        else:
            for i in range(8, 11):
                ordem[i].vagas = math.floor(vagasUniversais * ordem[i].vagas_perc * 2) if math.floor(vagasUniversais * ordem[i].vagas_perc*2) > 0 else 1
            for i in range(8, 11):
                vagasUniversais -= ordem[i].vagas
            if vagasUniversais > 0:
                for i in range(8, 11):
                    if vagasUniversais > 0:
                        ordem[i].vagas += 1
                        vagasUniversais -= 1
            if vagasUniversais < 0:
                for i in range(10, 7, -1):
                    if vagasUniversais < 0:
                        ordem[i].vagas -= 1
                        vagasUniversais += 1

        if vagasPublicas <= 8:
            sequencia = [9, 5, 7, 8, 3, 4, 6, 2]
            for i in range(8):
                for j in range(8):
                    if vagasPublicas > 0 and sequencia[i] == ordem[j].cota:
                        ordem[j].vagas += 1
                        vagasPublicas -= 1
        else:
            for i in range(8):
                ordem[i].vagas = math.floor(vagasPublicas * ordem[i].vagas_perc*2) if math.floor(vagasPublicas * ordem[i].vagas_perc*2) > 0 else 1
            for i in range(8):
                vagasPublicas -= ordem[i].vagas
            if vagasPublicas > 0:
                for i in range(8):
                    if vagasPublicas > 0:
                        ordem[i].vagas += 1
                        vagasPublicas -= 1
            if vagasPublicas < 0:
                for i in range(7, -1, -1):
                    if vagasPublicas < 0:
                        if ordem[i].vagas > 1:
                            ordem[i].vagas -= 1
                            vagasPublicas += 1
    for i in range(11):
        print(str(ordem[i].cota) + " " + str(ordem[i].vagas) + " - ", end='')
    print()
    return ordem


def fazerListasChamada(curso, candidatos, ordem):
    ordem_cham = destinarVagas(ordem, curso)
    ordem_cham.sort(key=lambda cota: cota.peso)

    onze_listas = []
    for i in range(1, 12, 1):
        onze_listas.append(listasParaCotas(curso, candidatos, i))

    data = []
    for i in range(len(onze_listas[0])):
        data.append(
            [onze_listas[0][i].codigo, onze_listas[0][i].nome, onze_listas[0][i].inscricao,
             onze_listas[0][i].posicao, onze_listas[0][i].matricula, onze_listas[0][i].chamada])
    df = pd.DataFrame(data, columns=['Código', 'Nome', 'Inscrição', 'Posição', 'Matrícula', 'Chamada'])
    df.to_excel(f'C:\\temp2\\{curso[1]}-{curso[0]}.xlsx', sheet_name = f'Cota-{1}', index=False)

    path = f'C:\\temp2\\{curso[1]}-{curso[0]}.xlsx'
    for i in range(1, 11):
        book = load_workbook(path)
        writer = pd.ExcelWriter(path, engine="openpyxl")
        writer.book = book

        data = []
        for j in range(len(onze_listas[i])):
            data.append([onze_listas[i][j].codigo, onze_listas[i][j].nome, onze_listas[i][j].inscricao,
                         onze_listas[i][j].posicao, onze_listas[i][j].matricula, onze_listas[i][j].chamada])
        df = pd.DataFrame(data, columns=['Código', 'Nome', 'Inscrição', 'Posição', 'Matrícula', 'Chamada'])
        df.to_excel(writer, sheet_name = f'Cota-{i+1}', index = False)
        writer.save()
        writer.close()



def listasParaCotas(curso, candidatos, cota):
    listaTemp = list(filter(lambda cand: cota in cand.cotas and cand.campus == curso[1] and cand.curso == curso[0],
                candidatos))
    listaTemp.sort(key=lambda cand: cand.posicao)
    return listaTemp