from openpyxl import load_workbook
import os
import pandas as pd
import re

from Classes import *

listasCursos = []
sequencia = [0, 9, 10, 8, 7, 6, 4, 5, 3, 2, 1]
primeirasChamadas = []


def recuperarListasIniciais():
    arquivos = os.listdir('C:\\temp2')
    for arq in arquivos:
        nome = re.split(r'[-.]+', arq)
        onzeListas = [Curso((nome[0], nome[1]))]
        listasIniciais = load_workbook('C:\\temp2\\' + arq)
        for i in range(1, 12, 1):
            listaCota = []
            planilha = listasIniciais.get_sheet_by_name(f'Cota-{i}')
            verificador = True
            count = 2
            while (verificador):
                if planilha[f'A{count}'].value == None:
                    verificador = False
                else:
                    listaCota.append(CandChamada(planilha[f'A{count}'].value, planilha[f'B{count}'].value,
                                                 planilha[f'D{count}'].value, planilha[f'E{count}'].value,
                                                 planilha[f'F{count}'].value))
                    count += 1
            onzeListas.append(listaCota)

        planilha = listasIniciais.get_sheet_by_name('Vagas')
        cursoVagas=[]
        for i in range(2, 13, 1):
            cursoVagas.append([planilha[f'A{i}'].value, planilha[f'B{i}'].value])
        onzeListas.append(cursoVagas)

        listasCursos.append(onzeListas)


def montarListaPrimeiraChamada(curso):
    chamada = []
    while True:
        montarLista(curso, chamada)
        if (curso[1][-1].chamada != 1 and any(vagas[1] != 0 for vagas in curso[12])):
            remanejarVagas(curso)
        else:
            break
    primeirasChamadas.append(chamada)


def montarLista(curso, chamada):
    for i in range(11):
        if curso[12][i][1] != 0:
            for cand in curso[sequencia[i]+1]:
                possivelChamado = list(filter(lambda x: x.codigo == cand.codigo, curso[1]))[0]
                if possivelChamado.chamada == 0:
                    possivelChamado.chamada = 1
                    chamada.append(possivelChamado)
                    curso[12][i][1] -= 1
                if curso[12][i][1] == 0:
                    break


def remanejarVagas(curso):
    for i in range(1, 11, 1):
        if curso[12][i][1] > 0:
            curso[12][i][1] -= 1
            curso[12][i-1][1] += 1


def fazerArquivosDeChamada():
    for chamada in range(len(primeirasChamadas)):
        data = []
        for c in range(len(primeirasChamadas[chamada])):
            data.append([primeirasChamadas[chamada][c].codigo,
                         primeirasChamadas[chamada][c].nome,
                         primeirasChamadas[chamada][c].posicao,
                         primeirasChamadas[chamada][c].matricula,
                         primeirasChamadas[chamada][c].chamada])
        df = pd.DataFrame(data, columns=['Código', 'Nome', 'Posição', 'Matrícula', 'Chamada'])
        df.to_excel(f'C:\\temp-PrimeiraChamada\\PC-{listasCursos[chamada][0].nome[0]}-{listasCursos[chamada][0].nome[1]}.xlsx', index=False)
