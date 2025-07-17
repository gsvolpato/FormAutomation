# MANUTENCAO

# Columns used in this project:
# - IDX: OS number in Sonner system
# - OS: Order Service number
# - Placa: Vehicle license plate
# - KM: Vehicle mileage
# - Tipo Serv: Service type (CORRETIVA or PREVENTIVA)
# - Item: Item description
# - Quantidade: Quantity
# - Valor Unit: Unit value
# - Categoria Item: Item category (MAO DE OBRA, PECAS, or LUBRIFICANTES)
# - Mao de Obra: Labor cost
# - Comentario: Comments
# - ANOEMPENHO: Budget year
# - EMPENHO: Budget number
# - NF: Invoice number
# - DATANF: Invoice date
# - Valor Total: Total value (Quantidade * Valor Unit)
# - Data OS: Order Service date
# - Status: Order Service status
# - Fornecedor: Supplier/Provider name
# - Serie NF: Invoice series

# 1 - Libraries
import time
import pandas as pd
import pyautogui
import sys
import keyboard

# Coordinates for all UI elements
COORDINATES = {
    # Main tabs
    "TAB_MAO_DE_OBRA": (348, 327),
    "TAB_PECAS": (435, 335),
    "TAB_LUBRIFICANTES": (518, 330),
    
    # Form fields
    "NUMERO_OS": (237, 238),
    "ITEM_FROTA": (243, 334),
    "SETA_TIPO": (499, 300),
    "KM": (771, 336),
    "AUTORIZADOR": (236, 441),
    
    # Fornecedor fields
    "FORNECEDOR_MAO_OBRA": (554, 379),
    "FORNECEDOR_PECAS_LUBES": (424, 371),
    
    # Discrimination buttons
    "DISCRIMINACAO_INCLUIR": (930, 599),
    "DISCRIMINACAO_ACEITAR": (930, 640),
    "DISCRIMINAR_ITEM": (914, 718),
    
    # Item actions
    "INCLUIR_ITEM": (702, 726),
    "INCLUIR_OS": (221, 156),

    "NF": (734, 567),
    "SERIE NF": (804, 564),   
    "DATA NF": (846, 567),    
    "ANO EMPENHO": (230, 557),
    "EMPENHO": (293, 557),    
    "COMENTARIO": (590, 494)
}

# Global pause variable
pause_execution = False

# Pause/Resume handler
def on_pause_press(e):
    global pause_execution
    if e.event_type == keyboard.KEY_DOWN:
        pause_execution = not pause_execution
        if pause_execution:
            print("\nPaused. Press 'Pause Break' again to resume.")
        else:
            print("\nResumed.")

# Register the pause key handler
keyboard.hook_key('pause', on_pause_press)

# Function to check if paused and wait if necessary
def check_pause():
    while pause_execution:
        time.sleep(0.1)

# 2 - Load the Excel file into a DataFrame
df = pd.read_excel('Filtered.xlsx')

# 3 - Confirmation Breaks
def confirm():
    pyautogui.hotkey('alt', 'tab')
    user_input = input("Press Enter to Continue or Q to quit: ").strip().upper()
    if user_input == "Q":
        print("Exiting...")
        exit()
    else:
        pass
    pyautogui.hotkey('alt', 'tab')

# 4 - Save Index
def save_manutencaoindex(idx):
    with open('manutencaoindex.txt', 'w') as file:
        file.write(str(idx))

# 5 - Load Index
def load_manutencaoindex():
    try:
        with open('manutencaoindex.txt', 'r') as file:
            return int(file.read())
    except FileNotFoundError:
        return 0

# 6 - Save Last OS
def save_last_os(os):
    with open('last_os.txt', 'w') as file:
        file.write(str(os))

# 7 - Load Last OS
def load_last_os():
    try:
        with open('last_os.txt', 'r') as file:
            return int(file.read())
    except FileNotFoundError:
        return 0

# 8 - Discriminação Mão de Obra
def mao_de_obra(row):
    print("____________________________________________________")
    check_pause() 
    pyautogui.moveTo(COORDINATES["TAB_MAO_DE_OBRA"]) # ABA MÃO DE OBRA
    pyautogui.click(COORDINATES["TAB_MAO_DE_OBRA"]) # ABA MÃO DE OBRA
    pyautogui.click(COORDINATES["TAB_MAO_DE_OBRA"]) # ABA MÃO DE OBRA
    time.sleep(0.25)
    check_pause() 
    pyautogui.moveTo(COORDINATES["FORNECEDOR_MAO_OBRA"]) # CAIXA FORNECEDOR MAO DE OBRA
    pyautogui.doubleClick(COORDINATES["FORNECEDOR_MAO_OBRA"]) # CAIXA FORNECEDOR MAO DE OBRA
    pyautogui.click(COORDINATES["FORNECEDOR_MAO_OBRA"]) # CAIXA FORNECEDOR MAO DE OBRA
    print("FORNECEDOR - MAO DE OBRA: 1148218")
    time.sleep(0.25)
    check_pause() 
    pyautogui.write('1148218'.upper(), interval=0.01)
    pyautogui.press('tab')
    time.sleep(0.25)
    check_pause() 
    pyautogui.moveTo(COORDINATES["DISCRIMINACAO_INCLUIR"]) # DISCRIMINAÇAO - INCLUIR
    pyautogui.click(COORDINATES["DISCRIMINACAO_INCLUIR"]) # DISCRIMINAÇAO - INCLUIR
    time.sleep(0.25)
    check_pause() 
    print("Categoria: Mao de Obra")
    pyautogui.write(('MAO DE OBRA').upper(), interval=0.01)
    time.sleep(0.25)
    check_pause() 
    pyautogui.press(['tab', 'tab'])
    time.sleep(0.25)
    check_pause() 
    print("Quantidade: 1")
    pyautogui.write('1'.upper(), interval=0.01)
    pyautogui.press('tab')
    time.sleep(0.25)
    check_pause() 
    print("Mao de Obra: ", row['Mao de Obra'])
    pyautogui.write(str(row['Mao de Obra']).replace('.', ',').upper())
    pyautogui.press('tab')
    time.sleep(0.25)
    check_pause() 
    pyautogui.moveTo(COORDINATES["DISCRIMINACAO_ACEITAR"]) # DISCRIMINAÇAO - ACEITAR
    pyautogui.click(COORDINATES["DISCRIMINACAO_ACEITAR"]) # DISCRIMINAÇAO - ACEITAR
    time.sleep(0.25)
    check_pause() 
    pyautogui.moveTo(COORDINATES["DISCRIMINACAO_ACEITAR"]) # DISCRIMINAÇAO - ACEITAR
    pyautogui.click(COORDINATES["DISCRIMINACAO_ACEITAR"]) # DISCRIMINAÇAO - ACEITAR
    time.sleep(0.25)

# 9 - Discriminação Peças e Lubrificantes
def produtos(row):
    print("____________________________________________________")
    check_pause() 
    pyautogui.doubleClick(COORDINATES["FORNECEDOR_PECAS_LUBES"]) # FORNECEDOR p PECAS e lubes
    pyautogui.doubleClick(COORDINATES["FORNECEDOR_PECAS_LUBES"]) # FORNECEDOR p PECAS e lubes
    time.sleep(0.25)
    check_pause() 
    print("FORNECEDOR - PECAS E LUBES: 1148218")
    pyautogui.write('1148218'.upper())
    time.sleep(0.25)
    check_pause() 
    pyautogui.moveTo(COORDINATES["DISCRIMINACAO_INCLUIR"]) # DISCRIMINAÇAO - INCLUIR
    pyautogui.click(COORDINATES["DISCRIMINACAO_INCLUIR"]) # DISCRIMINAÇAO - INCLUIR
    pyautogui.click(COORDINATES["DISCRIMINACAO_INCLUIR"]) # DISCRIMINAÇAO - INCLUIR
    print("ITEM: ", row['Item'])
    check_pause() 
    pyautogui.write(str(row['Item']).upper(), interval=0.01)
    time.sleep(0.25)
    check_pause() 
    pyautogui.press('tab')
    print("QUANTIDADE: ", row['Quantidade'])
    pyautogui.write(str(row['Quantidade']).upper(), interval=0.01)
    pyautogui.press('tab')
    check_pause() 
    print("VALOR UNIT: ", row['Valor Unit'])
    pyautogui.write(str(row['Valor Unit']).replace('.', ',').upper())
    pyautogui.press(['tab', 'tab'])
    check_pause() 
    if str(row['Categoria Item']) == "PECAS":
        print("Categoria: PECAS")
        pyautogui.write(('PE').upper(), interval=0.01)
        pyautogui.press(['down', 'up'])
    else:               
        print("Categoria: ", row['Categoria Item'])
        pyautogui.write(('LUBRIFICANTES').upper(), interval=0.01)
    pyautogui.press(['tab', 'tab','tab', 'tab'])
    check_pause() 

    instalacao = row['Mao de Obra']
    instalacao = row.get('Mao de Obra')
    if instalacao != '0,0':
        print("INSTALACAO: ", row['Mao de Obra'])
        mao_de_obra(row)
    else:
        pyautogui.moveTo(COORDINATES["DISCRIMINACAO_ACEITAR"]) # DISCRIMINAÇAO - ACEITAR
        pyautogui.click(COORDINATES["DISCRIMINACAO_ACEITAR"]) # DISCRIMINAÇAO - ACEITAR
        pyautogui.click(COORDINATES["DISCRIMINACAO_ACEITAR"]) # DISCRIMINAÇAO - ACEITAR
        time.sleep(0.25)

# 10 - Inclusão de Itens
def include_item(row):
    print("____________________________________________________")
    time.sleep(0.25)
    check_pause() 
    pyautogui.moveTo(COORDINATES["INCLUIR_ITEM"])  # INCLUIR ITEM
    pyautogui.click(COORDINATES["INCLUIR_ITEM"])  # INCLUIR ITEM
    pyautogui.click(COORDINATES["INCLUIR_ITEM"])  # INCLUIR ITEM
    time.sleep(0.25)
    check_pause() 
    pyautogui.press(['space', 'backspace'])
    print("Item: ", row['Item'])
    pyautogui.write(str(row['Item']).upper(), interval=0.01)
    time.sleep(0.25)
    check_pause() 
    pyautogui.press('tab')
    pyautogui.moveTo(COORDINATES["DISCRIMINAR_ITEM"]) # DISCRIMINAR ITEM
    time.sleep(1)
    check_pause() 
    pyautogui.doubleClick(COORDINATES["DISCRIMINAR_ITEM"]) # DISCRIMINAR ITEM
    time.sleep(1)
    check_pause() 
    
    tiposervico = row['Categoria Item']
    tiposervico = row.get('Categoria Item')

    instalacao = row['Mao de Obra']
    instalacao = row.get('Mao de Obra')

    if tiposervico == 'MAO DE OBRA':
        mao_de_obra(row)
    elif tiposervico == 'PECAS':
        time.sleep(0.25)
        check_pause() 
        pyautogui.moveTo(COORDINATES["TAB_PECAS"]) # ABA PECAS
        time.sleep(0.25)
        check_pause() 
        pyautogui.moveTo(COORDINATES["TAB_PECAS"]) # ABA PECAS
        pyautogui.click(COORDINATES["TAB_PECAS"]) # ABA PECAS
        time.sleep(0.25)
        check_pause() 
        produtos(row)
    elif tiposervico == 'LUBRIFICANTES':
        time.sleep(0.25)
        check_pause() 
        pyautogui.moveTo(COORDINATES["TAB_LUBRIFICANTES"]) # ABA LUBRIFICANTES
        time.sleep(0.25)
        check_pause() 
        pyautogui.moveTo(COORDINATES["TAB_LUBRIFICANTES"]) # ABA LUBRIFICANTES
        pyautogui.click(COORDINATES["TAB_LUBRIFICANTES"]) # ABA LUBRIFICANTES
        time.sleep(0.25)
        check_pause() 
        produtos(row)
        pyautogui.moveTo(COORDINATES["DISCRIMINACAO_INCLUIR"])  # DISCRIMINAÇAO - INCLUIR
        pyautogui.click(COORDINATES["DISCRIMINACAO_INCLUIR"])  # DISCRIMINAÇAO - INCLUIR
        time.sleep(0.25)

# 11 - Dados OS
def mains(row):
    print("____________________________________________________")
    time.sleep(0.25)
    check_pause() 
    pyautogui.moveTo(COORDINATES["NUMERO_OS"]) # NUMERO OS
    pyautogui.doubleClick(COORDINATES["NUMERO_OS"]) # NUMERO OS
    print("OS Sonner: ", row['IDX'])
    pyautogui.press('enter')
    time.sleep(0.25)
    check_pause() 
    pyautogui.press('enter')
    pyautogui.moveTo(COORDINATES["ITEM_FROTA"])  # ITEM DE FROTA
    pyautogui.doubleClick(COORDINATES["ITEM_FROTA"])  # ITEM DE FROTA
    time.sleep(0.25)
    check_pause() 
    print("Placa: ", row['Placa'])
    pyautogui.write(str(row['Placa']).upper(), interval=0.01)
    time.sleep(0.25)
    check_pause() 
    pyautogui.press(['tab' ,'space', 'space'])
    time.sleep(0.25)
    check_pause() 
    pyautogui.moveTo(COORDINATES["SETA_TIPO"]) # Seta Tipo
    pyautogui.doubleClick(COORDINATES["SETA_TIPO"]) # Seta Tipo
    pyautogui.doubleClick(COORDINATES["SETA_TIPO"]) # Seta Tipo
    time.sleep(0.25)
    check_pause() 
    tipo_serv_value = row["Tipo Serv"]
    if tipo_serv_value == "CORRETIVA":
        print("1 - MANUTENCAO CORRETIVA")
        pyautogui.write('1'.upper(), interval=0.01)
        pyautogui.press('down')
    elif tipo_serv_value == "PREVENTIVA":
        print("2 - MANUTENCAO PREVENTIVA")
        pyautogui.write('2'.upper(), interval=0.01)
        pyautogui.press('down')
    else:
        print("Unknown")
        sys.exit()

    time.sleep(0.25)
    check_pause() 
    pyautogui.moveTo(COORDINATES["KM"])  # KM
    pyautogui.click(COORDINATES["KM"])  # KM
    time.sleep(0.25)
    check_pause() 
    pyautogui.moveTo(COORDINATES["KM"])  # KM
    pyautogui.click(COORDINATES["KM"])  # KM
    time.sleep(0.25)
    check_pause() 
    print("KM: ", row['KM'])
    pyautogui.write(str(row['KM']).upper(), interval=0.01)
    time.sleep(0.25)
    check_pause() 
    pyautogui.moveTo(COORDINATES["AUTORIZADOR"])  # Autorizador
    pyautogui.click(COORDINATES["AUTORIZADOR"])  # Autorizador
    pyautogui.click(COORDINATES["AUTORIZADOR"])  # Autorizador
    time.sleep(0.25)
    check_pause() 
    print("AUTORIZADOR: ", row['AUTORIZADOR'])
    pyautogui.write(str(row['AUTORIZADOR']).upper(), interval=0.01)
    time.sleep(0.25)
    pyautogui.moveTo(COORDINATES["COMENTARIO"])  # Comentario
    pyautogui.click(COORDINATES["COMENTARIO"])  # Comentario
    pyautogui.click(COORDINATES["COMENTARIO"])  # Comentario
    time.sleep(0.25)
    check_pause() 
    print("Comentario: ", row['Comentario']) # Comentario
    pyautogui.write(f"{str(row['Comentario'])}".upper(), interval=0.01) # Comentario
    pyautogui.press('tab')
    time.sleep(0.25)
    check_pause() 
    pyautogui.moveTo(COORDINATES["ANO EMPENHO"])  # ANO EMPENHO
    pyautogui.click(COORDINATES["ANO EMPENHO"])  # ANO EMPENHO
    pyautogui.click(COORDINATES["ANO EMPENHO"])  # ANO EMPENHO
    print("ANO DO EMPENHO: ", row['ANOEMPENHO'])
    pyautogui.write(f"{str(row['ANOEMPENHO'])}".upper(), interval=0.01)
    time.sleep(0.25)
    check_pause() 
    pyautogui.press('tab')
    print("EMPENHO: ", row['EMPENHO'])
    pyautogui.write(f"{str(row['EMPENHO'])}".upper(), interval=0.01)
    time.sleep(0.25)
    check_pause() 
    pyautogui.press(['tab', 'tab', 'tab', 'tab',])
    time.sleep(0.25)
    check_pause() 
    print("NF: ", row['NF'])
    pyautogui.write(str(row['NF']).upper(), interval=0.01)
    time.sleep(0.25)
    check_pause() 
    pyautogui.press('tab')
    time.sleep(0.25)
    check_pause() 
    pyautogui.write('1'.upper(), interval=0.01)
    time.sleep(0.25)
    check_pause() 
    pyautogui.press('tab')
    time.sleep(0.25)
    check_pause() 
    print("DATA NF: ", '{:0>8}'.format(row['DATANF']))
    pyautogui.write('{:0>8}'.format(row['DATANF']).upper(), interval=0.01)  # Type data nota fiscal with leading zero
    time.sleep(0.25)
    check_pause() 
    pyautogui.press(['space', 'space', 'space'])


def click_at_positions():
    manutencaoindex = load_manutencaoindex()
    ultima_os = load_last_os()
    print("Starting from index:", manutencaoindex)
    print("Starting from OS ", ultima_os)
    print("Press 'Pause Break' key at any time to pause/resume execution")
    user_input = input("Press Enter to begin or Q to quit: ").strip().upper()
    print("____________________________________________________")
    if user_input == "Q":
        print("Exiting...")
        exit()
    pyautogui.hotkey('alt', 'tab')
    time.sleep(0.25)

    # Counter to track the index
    idx = manutencaoindex
    prev_os = ultima_os

    while idx < len(df):
        check_pause() 
        row = df.iloc[idx]
        current_os = row['OS']
        
        if current_os != prev_os:
            print("Iniciando Nova OS: ", current_os)
            mains(row)
            prev_os = current_os

        include_item(row)

        # Save the current state
        save_manutencaoindex(idx)
        save_last_os(current_os)
        
        # Check if the next row exists
        if idx + 1 < len(df):
            check_pause() 
            next_row = df.iloc[idx + 1]
            next_os = next_row['OS']
            if next_os != current_os:
                print("OS INCLUIDA")
                pyautogui.moveTo(COORDINATES["INCLUIR_OS"])  # INCLUIR OS
                pyautogui.click(COORDINATES["INCLUIR_OS"])  # INCLUIR OS
                pyautogui.click(COORDINATES["INCLUIR_OS"])  # INCLUIR OS
                pyautogui.press(['tab', 'space'])
                print("____________________________________________________")
                print("NOVA OS: ", next_os)

        idx += 1  # Move to the next row
        save_manutencaoindex(idx)

if __name__ == "__main__":
    try:
        click_at_positions()
    except KeyboardInterrupt:
        print("\nScript stopped by user")
    finally:
        keyboard.unhook_all()
        input("Press Enter to close...")
