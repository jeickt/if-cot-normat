from Funcoes import *

def main():

    candidatos = listarCandidatos()

    ordem = ordemChamada()

    for curso in cursosSel:
        listaChamada = fazerListasChamada(curso, candidatos, ordem)

if __name__ == "__main__":
    main()














