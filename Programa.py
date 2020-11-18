from Funcoes import *

def main():

    candidatos = listarCandidatos()

    for curso in cursosSel:
        ordem = ordemChamada()
        listaChamada = fazerListasChamada(curso, candidatos, ordem)

if __name__ == "__main__":
    main()














