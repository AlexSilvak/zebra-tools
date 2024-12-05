import win32print

# Configuração da impressora
printer_name = "ZDesigner"  # Substituir pelo nome correto da impressora
try:
    printer_handle = win32print.OpenPrinter(printer_name)
except Exception as e:
    print(f"Erro ao abrir a impressora: {e}")
    exit()

# Comandos ZPL
zpl_print_restaure_default = b"^XA^JUF^XZ"   # Restauração de fábrica
zpl_print_calibration = b"~JC^XA^XZ"         # Calibração
zpl_print_clear = b"~JA"                     # Limpeza da fila de impressão
zpl_print_teste = b"~WC"                     # Teste de impressão
zpl_print_clear_buffer= b"~JA"               # Limpeza do Buffer

# Menu de opções
while True:
    print("O nome da Impressora compartilhada deve ser ZDesigner")
    print("\nEscolha uma opção:")
    print("1 - Calibração")
    print("2 - Restauração de fábrica")
    print("3 - Limpeza da fila Impressão")
    print("4 - Teste de impressão")
    print("5 - Limpeza da Memoria de Impressão")
    print("6 - Escuridão")
    print("0 - Sair")

    try:
        option = int(input("Digite o número da opção: "))
        if option == 0:
            print("Finalizado.")
            break

        # Selecionar o comando baseado na opção
        if option == 1:
            zpl_command = zpl_print_calibration
        elif option == 2:
            zpl_command = zpl_print_restaure_default
        elif option == 3:
            zpl_command = zpl_print_clear
        elif option == 4:
            zpl_command = zpl_print_teste
        elif option == 5:
            zpl_print_clear_buffer 
        elif option == 6:
            # Solicitar o valor de escuridão ao usuário
            darkness_level = int(input("Digite o nível de escuridão (0 a 30): "))
            if 0 <= darkness_level <= 30:
                zpl_command = f"~SD{darkness_level}".encode('utf-8')
            else:
                print("Valor inválido. Digite um número entre 0 e 30.")
                continue       
        else:
            print("Opção inválida. Tente novamente.")
            continue

        # Envia o comando para a impressora
        job = win32print.StartDocPrinter(printer_handle, 1, ("Comando ZPL", None, "RAW"))
        win32print.StartPagePrinter(printer_handle)
        win32print.WritePrinter(printer_handle, zpl_command)
        win32print.EndPagePrinter(printer_handle)
        win32print.EndDocPrinter(printer_handle)
        print("Comando enviado com sucesso.")

    except ValueError:
        print("Por favor, digite um número válido.")
    except Exception as e:
        print(f"Erro ao enviar o comando para a impressora: {e}")

