print('Seja bem vindo a calculadora!\n')

while True:

    print('Escolha um número para seguir com sua operação!\n')
    print(
        '0 - Soma\n' \
        '1 - Subtração\n' \
        '2 - Multiplicação\n' \
        '3 - Divisão\n' \
        '4 - Exponenciação'
)

    escolha = -1

    while escolha not in [0, 1, 2, 3, 4]:
        try:
            escolha = int(input('Escolha um número:\n'))
            if escolha not in [0, 1, 2, 3, 4]:
                    print('Escolha incorreta, escolha um número entre as opções disponíveis\n')
        except ValueError:
            print('Digite apenas números inteiros!\n')

    print('\nOperação registrada!')
    print('Agora escolha dois números!')

    while True:
        try:
            numero1 = float(input('Escolha o primeiro número:\n'))
            numero2 = float(input('Escolha o segundo número:\n'))
            break
        except ValueError:
            print('Entrada inválida! Escolha apenas números!\n')

    print(f'Números escolhidos: {numero1} e {numero2}')

    if escolha == 0:
        print(f'A soma entre {numero1} e {numero2} é {numero1 + numero2}!')
    elif escolha == 1:
        print(f'A subtração entre {numero1} e {numero2} é {numero1 - numero2}!')
    elif escolha == 2:
        print(f'A multiplicação entre {numero1} e {numero2} é {numero1 * numero2}!')
    elif escolha == 3:
        while numero2 == 0:
            print("Erro: não é possível dividir por zero!")
            try:
                numero2 = float(input("Digite outro número para o divisor:\n"))
            except ValueError:
                print("Entrada inválida! Digite apenas números.\n")
                numero2 = 0  # força o while a continuar pedindo
        print(f'A divisão entre {numero1} e {numero2} é {numero1 / numero2}!')
    elif escolha == 4:
        print(f'A exponênciação entre {numero1} e {numero2} é {numero1 ** numero2}!')
   
    finalizar = input('Para sair da calculadora, digite q!\n')
    if finalizar == 'q':
        break