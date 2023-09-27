import tkinter as tk #biblioteca de exibição visual python
import socket #biblioteca para captura do nome de domínio
import getmac #biblioteca para captura do endereço físico da máquina
import wmi #biblioteca de acesso ao número de série vinculado a BIOS do notebook
import pymongo #biblioteca relacionado ao banco de dados MONGODB
import urllib.parse #biblioteca relacionado a necessidade de criptografia das informações de usuário e senha de acesso ao banco de dados
from pymongo.mongo_client import MongoClient #importação da propriedade MongoClient (acesso ao banco de dados e a coleção) de dentro da biblioteca pymongo
from pymongo.server_api import ServerApi #importação da comunicação via server com a API do banco de dados mongodb da biblioteca pymongo
import re #biblioteca utilizada para realizar uma busca dentro dos dados armazenados no programa para executar a função de simplificação do domínio
from tkinter import messagebox #importação do parametro de "caixa de mensagem" da biblioteca tkinter
from datetime import datetime #importação do parametro de data para gerar as informações de horário de verificação


# PROGRAMA PARA CAPTURA DE INFORMAÇÕES PERTINENTES DOS NOTEBOOKS NO CAMPUS BRAGANÇA PAULISTA


#captura o numero de série do notebook
def get_serial_number():
    try:
        c = wmi.WMI()
        serial_number = c.Win32_BIOS()[0].SerialNumber.strip()
        return serial_number
    except Exception as e:
        print(f"Erro ao obter o número de série: {e}")
        return "N/A"

#captura o mac address do notebook
def get_mac_addresses():
    mac_addresses = getmac.get_mac_address(
        interface="Ethernet", network_request=True)
    if isinstance(mac_addresses, str):
        return [mac_addresses]
    else:
        return mac_addresses

#captura o nome de domínio da máquina
def get_domain_name():
    domain_name = socket.getfqdn()
    return domain_name

# Obs: as informações são obtidas através de diretrizes de captura de informações semelhantes as utilizadas nos comandos executados no CMD do windows

# Simplicação do nome de domínio para melhor exibição no arquivo .csv gerado posteriomente
def simplificar_dominio(domain_name):
    # Procurar por padrões que se encaixem em "SP300LAB" seguido de 3 dígitos, hífen e mais 3 dígitos
    padrao = r"SP300LAB(\d{3})-(\d{3})"
    correspondencia = re.search(padrao, domain_name)

    if correspondencia:
        # Pegar o número do laboratório (três dígitos após "SP300LAB")
        numero_lab = int(correspondencia.group(1))
        # Pegar o número da máquina (três dígitos após o hífen)
        numero_maquina = int(correspondencia.group(2))
        # Extrair os dois últimos dígitos do número do laboratório e formatar como "LabXXX"
        lab_simplificado = f"Lab{str(numero_lab)[-3:]}"
        # Formatar o número da máquina como "YY"
        maquina_simplificada = str(numero_maquina).zfill(2)
        # Concatenar as partes formatadas em "LabXXX-YY"
        return f"{lab_simplificado}-{maquina_simplificada}"

    # Caso o padrão não seja encontrado, retornar o nome de domínio original
    return domain_name



# Configuração de acesso e armazenamento de dados no Banco de Dados (Não Relacional) MONGODB

def store_data():
    serial_number = result_label_serial['text']
    mac_addresses = result_label_mac['text']
    domain_name = result_label_domain['text']
    current_time = result_label_time['text']

    domain_name_simplificado = simplificar_dominio(domain_name)

    # Dados coletados pelo programa
    data = {
        'serial_number': serial_number,
        'mac_addresses': mac_addresses,
        'domain_name': domain_name_simplificado,
        'data_execucao': current_time
    }

    username = 'nicolasalvesalberti25'
    password = '@Ni820977'

    # Codifica o nome de usuário e a senha usando urllib.parse.quote_plus
    encoded_username = urllib.parse.quote_plus(username)
    encoded_password = urllib.parse.quote_plus(password)

    # Monta a string de conexão com os dados codificados
    connection_string = f"mongodb://{encoded_username}:{encoded_password}@ac-nm7ghql-shard-00-00.rfppyjj.mongodb.net:27017,ac-nm7ghql-shard-00-01.rfppyjj.mongodb.net:27017,ac-nm7ghql-shard-00-02.rfppyjj.mongodb.net:27017/?ssl=true&replicaSet=atlas-q4ubkv-shard-0&authSource=admin&retryWrites=true&w=majority"
    client = pymongo.MongoClient(connection_string)

    # Acessa o banco de dados e a coleção
    db = client['probooks']
    collection = db['lab05']

    # Insere os dados na coleção
    collection.insert_one(data)

    messagebox.showinfo("Sucesso", "Dados Coletados com Sucesso!")
    root.destroy()

   

#função principal da execução da funcionalidade do programa

def execute_program():
    serial_number = get_serial_number()
    mac_addresses = get_mac_addresses()
    domain_name = get_domain_name()
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # Formato: AAAA-MM-DD HH:MM:SS

    # Exibe os resultados formatados em rótulos
    result_label_serial.config(text=f"Serial Number: {serial_number}")
    result_label_mac.config(text=f"MAC Address: {', '.join(mac_addresses)}")
    result_label_domain.config(text=f"Domínio: {domain_name}")
    result_label_time.config(text=f"Hora da Verificação: {current_time}")


    # Habilita o botão para armazenar os dados (apenas após os dados serem exibidos)
    store_button.config(state=tk.NORMAL)



# Parte visual do programa utilizando a biblioteca visual TKINTER

# Cria a janela principal
root = tk.Tk()
root.title("IDENTIFICADOR DE NOTEBOOKS - UNIVERSIDADE SÃO FRANCISCO")

# Largura e altura da janela
largura_janela = 500
altura_janela = 400

# Obtém a largura e altura da tela
largura_tela = root.winfo_screenwidth()
altura_tela = root.winfo_screenheight()

# Calcula a posição central da janela
posicao_x = int((largura_tela / 2) - (largura_janela / 2))
posicao_y = int((altura_tela / 2) - (altura_janela / 2))

# Define a geometria da janela
root.geometry(f"{largura_janela}x{altura_janela}+{posicao_x}+{posicao_y}")

# Criação de um Label para exibir o texto grande no centro
header_label1 = tk.Label(
    root, text="IDENTIFICADOR DE NOTEBOOKS", font=("Arial", 14))
header_label1.pack(pady=5)

header_label2 = tk.Label(
    root, text="UNIVERSIDADE SÃO FRANCISCO", font=("Arial", 12))
header_label2.pack(pady=20)

# Criação de rótulos para exibir os resultados
result_label_serial = tk.Label(root, text="", font=("Arial", 12))
result_label_serial.pack(pady=5)
result_label_mac = tk.Label(root, text="", font=("Arial", 12))
result_label_mac.pack(pady=5)
result_label_domain = tk.Label(root, text="", font=("Arial", 12))
result_label_domain.pack(pady=5)
result_label_time = tk.Label(root, text="", font=("Arial", 8))
result_label_time.pack(pady=5)

# Criação de um botão para executar o programa
execute_button = tk.Button(
    root, text="Executar Verificação", command=execute_program)
execute_button.pack(pady=10)

# Criação do botão para armazenar os dados
store_button = tk.Button(root, text="Armazenar Dados", command=store_data, state=tk.DISABLED)
store_button.pack(pady=10)

# Referência ao criador do programa
creator_label = tk.Label(
    root, text="Criado e Desenvolvido por Nícolas Alberti", font=("Arial", 8))
creator_label.pack(side=tk.BOTTOM, pady=10)

# Inicia o loop principal da interface gráfica
root.mainloop()

#fim
