class Candidato:
    # classe para a primeira função de listagem dos candidatos em FuncoesIniciais.
    def __init__(self, codigo, nome, campus, nivel, curso, inscricao, cotas, pontuacao, posicao):
        self.codigo = codigo
        self.nome = nome
        self.campus = campus
        self.nivel = nivel
        self.curso = curso
        self.inscricao = inscricao
        self.cotas = cotas
        self.pontuacao = pontuacao if (self.nivel == "superior") else pontuacao * 2.5
        self.posicao = posicao

    def __str__(self):
        return self.codigo + "," + self.nome + "," + self.inscricao


class Cota:
    # classe para a realização dos ordenamentos em FuncoesIniciais.
    def __init__(self, cota, vagas_perc, peso):
        self.cota = cota
        self.vagas_perc = vagas_perc
        self.peso = peso
        self.vagas = 0

    def __str__(self):
        return [self.cota, self.vagas]

class Curso:
    # informações básicas do curso
    def __init__(self, nome):
        self.nome = nome
        self.vagas = []

class CandControle:
    # modelo para o acompanhamento principal dos resultados de cada chamada.
    def __init__(self, codigo, nome, inscricao, posicao, valido, matricula, chamada, ausencia, descRI, descPPI, descEP,
                 descPCD):
        self.codigo = codigo
        self.nome = nome
        self.inscricao = inscricao
        self.posicao = posicao
        self.valido = valido
        self.matricula = matricula
        self.chamada = chamada
        self.ausencia = ausencia
        self.descRI = descRI
        self.descPPI = descPPI
        self.descEP = descEP
        self.descPCD = descPCD

class CandResumido:
    # modelo resumido para os candidatos de cotas no acompanhamento principal e na elaboração das chamadas.
    def __init__(self, codigo, nome, inscricao, tipoVaga, chamada, valido):
        self.codigo = codigo
        self.nome = nome
        self.inscricao = inscricao
        self.tipoVaga = tipoVaga
        self.chamada = chamada
        self.valido = valido
