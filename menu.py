import os


def print_menu():  # Your menu design here
    print(30 * "-", "MENU", 30 * "-")

    print("1. Criar pastas")
    print("2. Vincular Supervisor")
    print("3. Vistoriar Obra")
    print("4. Programar Obra (Vistoria)")
    print("5. Programar Obra (Execução)")
    print("6. Finalizar Obra (GomMobile)")
    print("7. Extrair Sob & Status")
    print("8. Extrair Sob & Trabalho")
    print("9. Verificar Sob Energizada")
    print("10. Baixar arquivos DWG")
    print("11. Comparar arquivos DWG")
    print("11. Baixar arquivos SGD")

    print(67 * "-")


loop = True

while loop:  # Enquanto loop for True, o código continuará sendo executado
    print_menu()  # Exibe o menu
    choice = input("Enter your choice [1-5]: ")

    if choice == '1':

    elif choice == '2':

    elif choice == '3':

    elif choice == '4':

    elif choice == '5':

    elif choice == '6':

    elif choice == '7':

    elif choice == '8':

    elif choice == '9':

    elif choice == '10':

    elif choice == '11':

    elif choice == 'x':
        loop = False  # Esta opção fará com que a variável loop seja False, encerrando o programa
    else:
        input("Opção inválida, tente novamente.")