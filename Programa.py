import FuncoesIniciais as fi
import FuncoesPrimeiraChamada as fst
import FuncoesSegundaChamada as snd
import FuncoesTerceiraChamada as trd
import ChamadaPublica as cp


def main():

    print("""Qual é a etapa do Processo Seletivo?
                Opção 0 - Destinar vagas e obter listas para cada cota;
                Opção 1 - Confeccionar a chamada da etapa 1;
                Opção 2 - Confeccionar a chamada da etapa 2;
                Opção 3 - Confeccionar a chamada da etapa 3;
                Opção 4 - Realizar chamada pública.""")
    opcao = input()
    while opcao not in ["0", "1", "2", "3", "4"]:
        opcao = input("Opção inválida. Qual é a etapa do Processo Seletivo? ")

    if opcao == "0":
        candidatos = fi.listarCandidatos()
        for curso in fi.cursos:
            ordem = fi.ordemChamada()
            fi.fazerListasChamada(curso, candidatos, ordem)

    if opcao == "1":
        fst.recuperarListasIniciais()
        for curso in fst.listasCursos:
            fst.montarListaPrimeiraChamada(curso)
        fst.fazerArquivosDeChamada()
        fst.consolidarConferenciaPrincipal()

    if opcao == "2":
        snd.recuperarListasIniciais()
        snd.verificarDesclassificacoesEmCotas()
        for curso in snd.listasCursos:
            snd.montarListaSegundaChamada(curso)
        snd.fazerArquivosDeChamada()
        snd.consolidarConferenciaPrincipal()

    if opcao == "3":
        trd.recuperarListasIniciais()
        trd.verificarDesclassificacoesEmCotas()
        for curso in trd.listasCursos:
            trd.montarListaTerceiraChamada(curso)
        trd.fazerArquivosDeChamada()
        trd.consolidarConferenciaPrincipal()

    if opcao == "4":
        cp.recuperarListasIniciais()
        cp.verificarDesclassificacoesEmCotas()
        cp.inserirListasVazias()
        cp.comecarChamada()
        cp.arquivarChamadaPublica()
        cp.consolidarConferenciaPrincipal()


if __name__ == "__main__":
    main()
