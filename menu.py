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
    choice = input("Escolha uma opção [1-11]: ")

    if choice == '1':
        os.system('makeFolderSubfolder.py')
    elif choice == '2':
        os.system('VincSup.py')
    elif choice == '3':
        os.system('VistoriaSob.py')
    elif choice == '4':
        os.system('ProgramaSobVistoria.py')
    elif choice == '5':
        os.system('ProgramaSobExecucao.py')
    elif choice == '6':
        os.system('GOMMOBILE.py')
    elif choice == '7':
        os.system('ExtraiSobStatus.py')
    elif choice == '8':
        os.system('ExtraiSobTrabalho.py')
    elif choice == '9':
        os.system('VerificaSobEnergizada.py')
    elif choice == '10':
        os.system('DownloadDWG.py')
    elif choice == '11':
        os.system('DownloadSGD.py')
    elif choice == 'x':
        loop = False  # Esta opção fará com que a variável loop seja False, encerrando o programa
    else:
        input("Opção inválida, tente novamente.")