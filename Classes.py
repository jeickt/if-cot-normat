class Candidato:
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
        self.matricula = None
        self.chamada = 0

    def __str__(self):
        return self.codigo + "," + self.nome + "," + self.inscricao


class Vulnerabilidade:
    def __init__(self, fator, pontuacao_media):
        self.fator = fator
        self.pontuacao_media = pontuacao_media
        self.peso = 0

class Cota:
    def __init__(self, cota, vagas_perc, peso):
        self.cota = cota
        self.vagas_perc = vagas_perc
        self.peso = peso
        self.vagas = 0

    def __str__(self):
        return [self.cota, self.vagas]

class Curso:
    def __init__(self, nome):
        self.nome = nome
        self.vagas = []

class CandChamada:
    def __init__(self, codigo, nome, posicao, matricula, chamada):
        self.codigo = codigo
        self.nome = nome
        self.posicao = posicao
        self.matricula = matricula
        self.chamada = chamada