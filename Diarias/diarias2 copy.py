import pandas as pd
import pyautogui
import time
import unidecode  # Add this import at the top

def confirm():
    pyautogui.hotkey('alt', 'tab')
    input("Press Enter to continue...")
    time.sleep(0.5)
    pyautogui.hotkey('alt', 'tab')

# Load the Excel file into a DataFrame
df = pd.read_excel('imp2.xlsx', dtype=str)  # Ensure all data is read as strings

# Ensure specific columns are formatted correctly
df['Partida'] = df['Partida'].apply(lambda x: str(x).zfill(8))
df['Hora Partida'] = df['Hora Partida'].apply(lambda x: str(x).zfill(4))
df['Retorno'] = df['Retorno'].apply(lambda x: str(x).zfill(8))
df['Hora Retorno'] = df['Hora Retorno'].apply(lambda x: str(x).zfill(4))

def format_date(date):
    return str(date).replace('/', '-')  # Replace '/' with '-' in the date string

# Function to save and load the last processed index
def save_impIndex(index):
    with open('impIndex.txt', 'w') as file:
        file.write(str(index))

def load_impIndex():
    try:
        with open('impIndex.txt', 'r') as file:
            return int(file.read())
    except FileNotFoundError:
        return 0

def click_at_positions():
    impIndex = load_impIndex()
    print("Starting from index:", impIndex)
    input("Press Enter to continue...")
    time.sleep(0.5)
    print("Pressing Alt+Tab...")
    pyautogui.hotkey('alt', 'tab')
    time.sleep(0.5) 
    
    # Counter to track the index
    idx = impIndex
    
    # Iterate over the rows of the DataFrame
    for _, row in df.iloc[impIndex:].iterrows():
        pyautogui.click((973, 276)) # Nova Diaria
        pyautogui.click((412, 153)) # Nome
        
        # Remove accents from Nome before writing
        nome = unidecode.unidecode(str(row['Nome']))
        pyautogui.write(nome, interval=0.01)
        print("Nome:", str(row['Nome']))
        
        pyautogui.press('tab')
        time.sleep(0.5)
        pyautogui.write('JO')
        time.sleep(0.5)
        pyautogui.press('tab')
        
        # Remove accents from Destino before writing
        destino = unidecode.unidecode(str(row['Destino']))
        pyautogui.write(destino, interval=0.01)
        print("Destino: ", str(row['Destino']))
        
        pyautogui.press('tab')
        pyautogui.write('S') # Tp Hospedagem: Sem Hospedagem
        pyautogui.press(['tab', 'tab'])
        pyautogui.write(row['Partida'], interval=0.01)
        print("Partida: ", row['Partida'])
        pyautogui.press('tab')
        pyautogui.write(row['Hora Partida'], interval=0.01)
        print("Hora Partida: ", row['Hora Partida'])
        pyautogui.press('tab')
        pyautogui.write(row['Retorno'], interval=0.01)
        print("Retorno: ", row['Retorno'])
        pyautogui.press('tab')
        pyautogui.write(row['Hora Retorno'], interval=0.01)
        print("Hora Retorno: ", row['Hora Retorno'])
        pyautogui.press('tab')

        #pyautogui.write(str(row['Transporte']), interval=0.01) # Meio de Transporte
        #print("Transporte: ", str(row['Transporte']))

        pyautogui.write('V') # Veiculo Oficial
        pyautogui.press('tab')
        pyautogui.write(str(row['Veiculo']), interval=0.01)
        print("Veiculo: ", str(row['Veiculo']))
        pyautogui.press(['tab', 'tab'])
        pyautogui.write('S') # Tanque Cheio
        pyautogui.press(['tab', 'tab', 'tab'])

        #pyautogui.write(str(row['Passagem']), interval=0.01)
        #print("Passagem: ", str(row['Passagem']))

        pyautogui.write('0') # Passagem
        pyautogui.press(['tab', 'tab', 'tab', 'tab'])
        
        pyautogui.write(str(row['Pedagio']).replace('.', ','), interval=0.01)
        print("Pedagio: ", str(row['Pedagio']))
        pyautogui.press(['tab', 'tab', 'tab'])

        pyautogui.write('V') # Alimentação Inteira -> Valor Alimentação
        pyautogui.press('tab')

        if row['MeiaAlim'] == '1':
            pyautogui.write('V', interval=0.01) # Valor 1/2 Alimentação
            print("MeiaAlim: ", str(row['MeiaAlim']))
        else:
            pyautogui.write(str(int(row['MeiaAlim'])), interval=0.01)
            print("MeiaAlim: ", str(int(row['MeiaAlim'])))

        pyautogui.press('tab')

        if row['MeiaHosp'] == 1:
            pyautogui.write('V', interval=0.01) # Valor 1/2 Hospedagem
            print("Meia Hospedagem: ", str(row['MeiaHosp']))
        else:
            pyautogui.write(str(int(row['MeiaHosp'])), interval=0.01)
            print("Meia Hospedagem: ", str(int(row['MeiaHosp'])))

        pyautogui.press(['tab', 'tab'])
        pyautogui.write(str(int(row['Alim'])), interval=0.01)
        print("Alimentacao: ", str(row['Alim']))
        pyautogui.press('tab')
        pyautogui.write(str(int(row['Hosp'])), interval=0.01)
        print("Hospedagem: ", str(row['Hosp']))
        pyautogui.press('tab')
        
        # Remove accents from Finalidade before writing
        finalidade = unidecode.unidecode(str(row['Finalidade']))
        pyautogui.write(finalidade, interval=0.01)
        print("Finalidade: ", str(row['Finalidade']))
        
        pyautogui.press('tab')
        
        obs_value = row['Obs']
        # Check if the value is not NaN and not an empty string
        if pd.notna(obs_value) and obs_value != '':
            # Remove accents from Obs before writing
            obs = unidecode.unidecode(str(obs_value))
            pyautogui.write(obs, interval=0.01)

        time.sleep(0.5)
        pyautogui.click((983, 223)) # SALVAR
        time.sleep(0.5)
        pyautogui.click((962, 406))  # FSD
        time.sleep(1)
        pyautogui.click((944, 94))  # PDF
        time.sleep(1)
        pyautogui.write(str(row['NomeFSD']), interval=0.01)
        print("NomeFSD: ", str(row['NomeFSD']))
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.press('esc')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(0.5)
        pyautogui.click((1124, 108))  # FECHAR VISUALIZAÇÃO
        time.sleep(0.5)
        pyautogui.click((962, 455))  # RELATORIO
        time.sleep(0.5)
        pyautogui.click((577, 472))  # DESCRIÇAO
        
        # Remove accents from Finalidade before writing in report
        finalidade = unidecode.unidecode(str(row['Finalidade']))
        pyautogui.write(finalidade, interval=0.01)
        print("Finalidade: ", str(row['Finalidade']))
        
        pyautogui.click((930, 479))  # DOCUMENTOS
        #pyautogui.write(str(row['DOCS RV']), interval=0.01)
        #print("DOCS RV: ", str(row['DOCS RV']))
        pyautogui.write('Anexados ao relatorio.')
        pyautogui.click((326, 570)) # OUTRAS CONSIDERAÇOES
        #pyautogui.write(str(row['Obs']), interval=0.01)
        
        obs_value = row['Obs']
        if obs_value == 'nan':
            pyautogui.write('')

        pyautogui.click((1039, 266))  # SALVAR RV
        time.sleep(0.5)
        pyautogui.click((1043, 644))  # IMPRIMIR RV
        time.sleep(0.5)
        pyautogui.click((1110, 49))  # FECHAR RV
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(0.5)
        pyautogui.click((944, 94))  # PDF
        time.sleep(0.5)
        pyautogui.write(str(row['NomeRV']), interval=0.01)
        print("NomeRV: ", str(row['NomeRV']))
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.press('esc')
        time.sleep(0.5)
        pyautogui.click((1124, 108))  # FECHAR VISUALIZAÇÃO
        pyautogui.click((1124, 108))  # FECHAR VISUALIZAÇÃO
        save_impIndex(idx + 1)  # Save the index of the last processed row
        idx += 1
        time.sleep(1)

if __name__ == "__main__":
    click_at_positions()
    pyautogui.hotkey('alt', 'tab')    
    input("Press Enter to close...")  # Wait for user input to close the prompt
