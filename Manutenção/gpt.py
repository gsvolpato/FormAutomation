# MANUTENCAO

# 1 - Libraries
import time
import pandas as pd
import pyautogui
import sys
import keyboard

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
df = pd.read_excel('filtered02864324-cleaned.xlsx')

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
    check_pause() # Check if paused
    pyautogui.moveTo((348, 327)) # ABA MÃO DE OBRA
    pyautogui.click((348, 327)) # ABA MÃO DE OBRA
    time.sleep(0.5)
    check_pause() # Check if paused
    pyautogui.moveTo((554, 379)) # CAIXA FORNECEDOR MAO DE OBRA
    pyautogui.click((554, 379)) # CAIXA FORNECEDOR MAO DE OBRA
    pyautogui.click((554, 379)) # CAIXA FORNECEDOR MAO DE OBRA
    print("FORNECEDOR - MAO DE OBRA: 1148218")
    time.sleep(0.5)
    check_pause() # Check if paused
    pyautogui.write('1148218'.upper(), interval=0.01)
    pyautogui.press('tab')
    time.sleep(0.5)
    check_pause() # Check if paused
    pyautogui.moveTo((930, 599)) # DISCRIMINAÇAO - INCLUIR
    pyautogui.click((930, 599)) # DISCRIMINAÇAO - INCLUIR
    time.sleep(0.5)
    check_pause() # Check if paused
    print("Categoria: Mao de Obra")
    pyautogui.write(('MAO DE OBRA').upper(), interval=0.01)
    time.sleep(0.5)
    check_pause() # Check if paused
    pyautogui.press(['tab', 'tab'])
    time.sleep(0.5)
    check_pause() # Check if paused
    print("Quantidade: 1")
    pyautogui.write('1'.upper(), interval=0.01)
    pyautogui.press('tab')
    time.sleep(0.5)
    check_pause() # Check if paused
    print("Mao de Obra: ", row['Mao de Obra'])
    pyautogui.write(str(row['Mao de Obra']).replace('.', ',').upper())
    pyautogui.press('tab')
    time.sleep(0.5)
    check_pause() # Check if paused
    pyautogui.moveTo((930, 640)) # DISCRIMINAÇAO - ACEITAR
    #pyautogui.doubleClick((930, 640)) # DISCRIMINAÇAO - ACEITAR
    pyautogui.click((930, 640)) # DISCRIMINAÇAO - ACEITAR
    time.sleep(0.5)
    check_pause() # Check if paused
    pyautogui.click((930, 640)) # DISCRIMINAÇAO - ACEITAR
    #print("DISCRIMINAÇAO - ACEITAR")
    time.sleep(0.5)

# 9 - Discriminação Peças e Lubrificantes
def produtos(row):
    print("____________________________________________________")
    check_pause() # Check if paused
    pyautogui.doubleClick((424, 371)) # FORNECEDOR p PECAS e lubes
    pyautogui.doubleClick((424, 371)) # FORNECEDOR p PECAS e lubes
    time.sleep(0.5)
    check_pause() # Check if paused
    print("FORNECEDOR - PECAS E LUBES: 1148218")
    pyautogui.write('1148218'.upper())
    time.sleep(0.5)
    check_pause() # Check if paused
    #pyautogui.press('enter')
    pyautogui.moveTo((930, 599)) # DISCRIMINAÇAO - INCLUIR
    pyautogui.click((930, 599)) # DISCRIMINAÇAO - INCLUIR
    pyautogui.click((930, 599)) # DISCRIMINAÇAO - INCLUIR
    print("ITEM: ", row['Item'])
    check_pause() # Check if paused
    pyautogui.write(str(row['Item']).upper(), interval=0.01)
    time.sleep(0.5)
    check_pause() # Check if paused
    pyautogui.press('tab')
    print("QUANTIDADE: ", row['Quantidade'])
    pyautogui.write(str(row['Quantidade']).upper(), interval=0.01)
    pyautogui.press('tab')
    check_pause() # Check if paused
    print("VALOR UNIT: ", row['Valor Unit'])
    pyautogui.write(str(row['Valor Unit']).replace('.', ',').upper())
    #pyautogui.write(f"{float(row['Valor Unit']):.2f}".replace('.', ','))
    pyautogui.press(['tab', 'tab'])
    check_pause() # Check if paused
    if str(row['Categoria Item']) == "PECAS":
        print("Categoria: PECAS")
        pyautogui.write(('PE').upper(), interval=0.01)
        pyautogui.press(['down', 'up'])
    else:               
        print("Categoria: ", row['Categoria Item'])
        pyautogui.write(('LUBRIFICANTES').upper(), interval=0.01)
    pyautogui.press(['tab', 'tab','tab', 'tab'])
    check_pause() # Check if paused

    instalacao = row['Mao de Obra']
    instalacao = row.get('Mao de Obra')
    if instalacao != '0,0':
        print("INSTALACAO: ", row['Mao de Obra'])
        mao_de_obra(row)
    else:
        #sys.exit()
        pyautogui.moveTo((930, 640)) # DISCRIMINAÇAO - ACEITAR
        pyautogui.click((930, 640)) # DISCRIMINAÇAO - ACEITAR
        pyautogui.click((930, 640)) # DISCRIMINAÇAO - ACEITAR
        time.sleep(0.5)

# 10 - Inclusão de Itens
def include_item(row):
    print("____________________________________________________")
    time.sleep(0.5)
    check_pause() # Check if paused
    pyautogui.moveTo((702, 726))  # INCLUIR ITEM
    pyautogui.click((702, 726))  # INCLUIR ITEM
    pyautogui.click((702, 726))  # INCLUIR ITEM
    time.sleep(0.5)
    check_pause() # Check if paused
    pyautogui.press(['space', 'backspace'])
    print("Item: ", row['Item'])
    pyautogui.write(str(row['Item']).upper(), interval=0.01)
    time.sleep(0.5)
    check_pause() # Check if paused
    pyautogui.press('tab')
    pyautogui.moveTo((914, 718)) # DISCRIMINAR ITEM
    time.sleep(1)
    check_pause() # Check if paused
    pyautogui.doubleClick((914, 718)) # DISCRIMINAR ITEM
    time.sleep(1)
    check_pause() # Check if paused
    
    tiposervico = row['Categoria Item']
    tiposervico = row.get('Categoria Item')

    instalacao = row['Mao de Obra']
    instalacao = row.get('Mao de Obra')

    if tiposervico == 'MAO DE OBRA':
        mao_de_obra(row)
    elif tiposervico == 'PECAS':
        time.sleep(0.5)
        check_pause() # Check if paused
        pyautogui.moveTo((435, 335)) # ABA PECAS
        time.sleep(0.5)
        check_pause() # Check if paused
        pyautogui.moveTo((435, 335)) # ABA PECAS
        pyautogui.click((435, 335)) # ABA PECAS
        time.sleep(0.5)
        check_pause() # Check if paused
        produtos(row)
    elif tiposervico == 'LUBRIFICANTES':
        time.sleep(0.5)
        check_pause() # Check if paused
        pyautogui.moveTo((518, 330)) # ABA LUBRIFICANTES
        time.sleep(0.5)
        check_pause() # Check if paused
        pyautogui.moveTo((518, 330)) # ABA LUBRIFICANTES
        pyautogui.click((518, 330)) # ABA LUBRIFICANTES
        time.sleep(0.5)
        check_pause() # Check if paused
        produtos(row)
        pyautogui.moveTo((930, 599))  # DISCRIMINAÇAO - INCLUIR
        pyautogui.click((930, 599))  # DISCRIMINAÇAO - INCLUIR
        time.sleep(0.5)

# 11 - Dados OS
def mains(row):
    print("____________________________________________________")
    time.sleep(0.5)
    check_pause() # Check if paused
    pyautogui.moveTo((237, 238)) # NUMERO OS
    pyautogui.doubleClick((237, 238)) # NUMERO OS
    print("OS Sonner: ", row['IDX'])
    pyautogui.press('enter')
    time.sleep(0.5)
    check_pause() # Check if paused
    pyautogui.press('enter')
    pyautogui.moveTo((243, 334))  # ITEM DE FROTA
    pyautogui.doubleClick((243, 334))  # ITEM DE FROTA
    time.sleep(0.5)
    check_pause() # Check if paused
    print("Placa: ", row['Placa'])
    pyautogui.write(str(row['Placa']).upper(), interval=0.01)
    time.sleep(0.5)
    check_pause() # Check if paused
    pyautogui.press(['tab' ,'space', 'space'])
    time.sleep(0.5)
    check_pause() # Check if paused
    pyautogui.moveTo((499, 300)) # Seta Tipo
    pyautogui.doubleClick((499, 300)) # Seta Tipo
    pyautogui.doubleClick((499, 300)) # Seta Tipo
    time.sleep(0.5)
    check_pause() # Check if paused
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

    time.sleep(0.5)
    check_pause() # Check if paused
    pyautogui.click((771, 336))  # KM
    time.sleep(0.5)
    check_pause() # Check if paused
    pyautogui.click((771, 336))  # KM
    time.sleep(0.5)
    check_pause() # Check if paused
    print("KM: ", row['KM'])
    pyautogui.write(str(row['KM']).upper(), interval=0.15)
    time.sleep(0.5)
    check_pause() # Check if paused
    pyautogui.click((236, 441))  # Autorizador
    pyautogui.click((236, 441))  # Autorizador
    time.sleep(0.5)
    check_pause() # Check if paused
    print("AUTORIZADOR: 138208272")
    pyautogui.write('138208272'.upper(), interval=0.01)
    time.sleep(0.5)
    check_pause() # Check if paused
    pyautogui.press('tab')
    time.sleep(0.5)
    check_pause() # Check if paused
    pyautogui.press('space')
    time.sleep(0.5)
    check_pause() # Check if paused
    #pyautogui.write(str(row['Descrição']), interval=0.01)
    pyautogui.press('tab')
    time.sleep(0.5)
    check_pause() # Check if paused
    print("Comentario: ", row['Comentario'])
    pyautogui.write(f"{str(row['Comentario'])}".upper(), interval=0.01)
    #confirm()
    pyautogui.press('tab')
    time.sleep(0.5)
    check_pause() # Check if paused
    print("ANO DO EMPENHO: 2024")
    pyautogui.write('2024'.upper(), interval=0.01)
    time.sleep(0.5)
    check_pause() # Check if paused
    pyautogui.press('tab')
    print("EMPENHO: 689")
    pyautogui.write('689'.upper(), interval=0.01)
    time.sleep(0.5)
    check_pause() # Check if paused
    pyautogui.press(['tab', 'tab', 'tab', 'tab',])
    time.sleep(0.5)
    check_pause() # Check if paused
    print("NF: ", row['NF'])
    pyautogui.write(str(row['NF']).upper(), interval=0.01)
    time.sleep(0.5)
    check_pause() # Check if paused
    pyautogui.press('tab')
    time.sleep(0.5)
    check_pause() # Check if paused
    pyautogui.write('1'.upper(), interval=0.01)
    time.sleep(0.5)
    check_pause() # Check if paused
    pyautogui.press('tab')
    time.sleep(0.5)
    check_pause() # Check if paused
    print("DATA NF: ", '{:0>8}'.format(row['DATANF']))
    pyautogui.write('{:0>8}'.format(row['DATANF']).upper(), interval=0.01)  # Type data nota fiscal with leading zero
    time.sleep(0.5)
    check_pause() # Check if paused
    pyautogui.press(['space', 'space', 'space'])
    #confirm()


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
    time.sleep(0.5)

    # Counter to track the index
    idx = manutencaoindex
    prev_os = ultima_os

    while idx < len(df):
        check_pause() # Check if paused
        row = df.iloc[idx]
        current_os = row['OS']
        
        if current_os != prev_os:
            print("Iniciando Nova OS: ", current_os)
            mains(row)
            prev_os = current_os

        include_item(row)
        #confirm()

        # Save the current state
        save_manutencaoindex(idx)
        save_last_os(current_os)
        
        # Check if the next row exists
        if idx + 1 < len(df):
            check_pause() # Check if paused
            next_row = df.iloc[idx + 1]
            next_os = next_row['OS']
            if next_os != current_os:
                print("OS INCLUIDA")
                pyautogui.click((221, 156))  # INCLUIR OS
                pyautogui.click((221, 156))  # INCLUIR OS
                pyautogui.press(['tab', 'space'])
                print("____________________________________________________")
                print("NOVA OS: ", next_os)
                #confirm()

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
