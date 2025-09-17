import os

def limpar_tela():
  os.system('cls' if os.name == 'nt' else 'clear')

carros = [
    ('Chevrolet Tracker', 120),
    ('Chevrolet Onix', 90),
    ('Chevrolet Spin', 120),
    ('Hyundai HB20', 85),
    ('Hyundai Tucson', 120),
    ('Fiat Uno', 60),
    ('Fiat Mobi', 70),
    ('Fiat Pulse', 130)
]

alugados = []

def mostrar_lista_de_carros(lista_de_carros):
  if not lista_de_carros:
    print('Nenhum carros disponível.')
  else:
    for i, car in enumerate(lista_de_carros):
      print(f'{i} - {car[0]} - R$ {car[1]} /dia.')

while True:
  limpar_tela()
  print('=' * 50)
  print('Bem-vindo a locadora de carros!')
  print('=' * 50)
  print('O que você deseja fazer?.')
  print('0 - Mostrar portifólio')
  print('1 - Alugar Carro')
  print('2 - Devolver Carro')
  print('3 - Sair')

  try:
    opcao = int(input('\nDigite a opção: '))
  except ValueError:
    print('Digite apenas números válidos!')
    input('Pressione Enter para continuar...')
    continue

  if opcao == 0:
    limpar_tela()
    print('PORTFÓLIO DE CARROS DISPONÍVEIS:\n')
    mostrar_lista_de_carros(carros)
    input('\nPressione Enter para voltar ao menu...')

  elif opcao == 1:
    limpar_tela()
    print('CARROS DISPONÍVEIS PARA ALUGAR:\n')
    mostrar_lista_de_carros(carros)

    if carros:
      try:
        escolha = int(input('\nDigite o número do carro que deseja alugar: '))
        alugados.append(carros.pop(escolha))
        print('Carro alugado com sucesso!')
      except (ValueError, IndexError):
        print('Escolha inválida!')
    input('\nPressione Enter para voltar ao menu...')

  elif opcao == 2:
    limpar_tela()
    print('CARROS ALUGADOS:\n')
    mostrar_lista_de_carros(alugados)

    if alugados:
      try:
        escolha = int(input('\nDigite o número do carro que deseja devolver: '))
        carros.append(alugados.pop(escolha))
        print('Carro devolvido com sucesso!')
      except (ValueError, IndexError):
        print('Escolha inválida!')
    input('\nPressione Enter para voltar ao menu...')

  elif opcao == 3:
    print('Obrigado por usar a locadora de carros!')
    break
  else:
    print('Opção inválida!')
    input('Pressione Enter para continuar...')


