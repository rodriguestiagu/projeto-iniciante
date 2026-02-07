import random

class Academia:
    def __init__(self):
        self.halteres = [i for i in range(10, 36) if i % 2 == 0]
        self.porta_halteres = {}
        self.reiniciar_o_dia()
    
    def reiniciar_o_dia(self):
        self.porta_halteres = {i: i for i in self.halteres}

    def listar_halteres(self):
        return [i for i in self.porta_halteres.values() if i != 0]
    
    def listar_espacos(self):
        return [i for i, j in self.porta_halteres.items() if j == 0]

    def pegar_halteres(self, peso):
        halt_pos = list(self.porta_halteres.values()).index(peso)
        key_halt = list(self.porta_halteres.keys())[halt_pos]
        self.porta_halteres[key_halt] = 0
        return peso
    
    def devolver_halter(self, pos, peso):
        self.porta_halteres[pos] = peso

    def calcular_caos(self):
        num_caos = [i for i, j in self.porta_halteres.items() if i != j]
        return len(num_caos) / len(self.porta_halteres)

class Usuario:
    def __init__(self, tipo, academia):
        self.tipo = tipo
        self.academia = academia
        self.peso = 0

    def iniciar_treino(self):
        lista_peso = self.academia.listar_halteres()
        self.peso = random.choice(lista_peso)
        self.academia.pegar_halteres(self.peso)

    def finalizar_treino(self):
        espacos = self.academia.listar_espacos()
        
        if self.tipo == 1:
            if self.peso in espacos:
                self.academia.devolver_halter(self.peso, self.peso)
            else:
                if espacos:  # Verificação adicional
                    pos = random.choice(espacos)
                    self.academia.devolver_halter(pos, self.peso)
                else:
                    print("Nenhum espaço disponível para tipo 1!")

        if self.tipo == 2:
            if espacos:  # Verificação adicional
                pos = random.choice(espacos)
                self.academia.devolver_halter(pos, self.peso)
            else:
                print("Nenhum espaço disponível para tipo 2!")
        self.peso = 0

# Código principal
academia = Academia()

usuarios = [Usuario(1, academia) for i in range(10)]
usuarios += [Usuario(2, academia) for i in range(1)]
random.shuffle(usuarios)

list_chaos = []

for k in range(50):
    academia.reiniciar_o_dia()
    for i in range(10):
        random.shuffle(usuarios)
        for user in usuarios:
            user.iniciar_treino()
        for user in usuarios:
            user.finalizar_treino()
    list_chaos.append(academia.calcular_caos())  # Use append em vez de += para listas

# Parte de plotagem com tratamento de erro
try:
    import seaborn as sns
    import matplotlib.pyplot as plt
    sns.displot(list_chaos)
    plt.show()
except ImportError:
    print("Seaborn não instalado. Usando matplotlib como alternativa...")
    import matplotlib.pyplot as plt
    plt.hist(list_chaos, bins=10, alpha=0.7)
    plt.xlabel('Nível de Caos')
    plt.ylabel('Frequência')
    plt.title('Distribuição do Caos na Academia')
    plt.show()