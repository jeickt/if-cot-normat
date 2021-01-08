from FuncoesIniciais import *
from FuncoesPrimeiraChamada import *


def main():

    print("""Qual é a etapa do Processo Seletivo?
                Opção 0 - Destinar vagas e obter listas para cada cota;
                Opção 1 - Confeccionar a chamada da etapa 1;
                Opção 2 - Confeccionar a chamada da etapa 2;
                Opção 3 - Confeccionar a chamada da etapa 3.""")
    opcao = int(input())

    if opcao == 0:
        candidatos = listarCandidatos()
        for curso in cursos:
            ordem = ordemChamada()
            fazerListasChamada(curso, candidatos, ordem)

    if opcao == 1:
        recuperarListasIniciais()
        for curso in listasCursos:
            montarListaPrimeiraChamada(curso)
        fazerArquivosDeChamada()

    if opcao == 2:
        pass

    if opcao == 3:
        pass

    if opcao == 2:
        pass


if __name__ == "__main__":
    main()














