    
import logging
import pandas as pd
from transform import Transform
from extract import Extract
#import pyautogui
import win32com.client as win32
import pythoncom
import time
import logging

logging.basicConfig(
    level=logging.INFO,  # Exibe mensagens a partir de INFO
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.StreamHandler()  # Garante logs no console
    ]
)

today = pd.Timestamp.today()
yesterday_format = (today-pd.Timedelta(days=1)).date().strftime('%d/%m/%Y')

class Load:
    def to_outlook(tabela_df, df_ativ):

            
            pythoncom.CoInitialize()

            # Garantir que a coluna de data está no formato correto
            df_ativ['data_ativacao'] = pd.to_datetime(df_ativ['data_ativacao'])

            # Converte DataFrame para tabela HTML com índice
            tabela_html = tabela_df.to_html(classes='tabela', index=True)  # Agora inclui o índice

            # CSS embutido no HTML
            css_style = """
                <style>
                    /* Estiliza o cabeçalho das colunas, o índice, a última coluna e a última linha */
                    .tabela thead th, 
                    .tabela tbody th, 
                    .tabela td:last-child, 
                    .tabela tbody tr:last-child td {
                        background-color: #f0f0f0; /* Cinza claro */
                        color: black; /* Fonte preta */
                        font-weight: bold; /* Negrito */
                        text-align: center; /* Centraliza os cabeçalhos e as células */
                        padding: 8px; /* Adiciona espaçamento interno */
                    }

                    /* Estiliza todas as outras células */
                    .tabela td {
                        background-color: white; /* Fundo branco */
                        color: black; /* Fonte preta */
                        text-align: center; /* Centraliza os valores */
                        padding: 8px; /* Adiciona espaçamento interno */
                    }

                    /* Estiliza a célula do total geral (interseção da última linha e última coluna) */
                    .tabela tbody tr:last-child td:last-child {
                        background-color: #404040; /* Cinza escuro */
                        color: white; /* Texto branco */
                        font-weight: bold;
                    }

                    /* Adiciona uma borda superior preta fina à última linha da tabela */
                    .tabela tbody tr:last-child {
                        border-top: 1px solid black; /* Borda superior preta fina */
                    }

                    /* Adiciona bordas finas cinza entre as células */
                    .tabela th, 
                    .tabela td {
                        border: 1px solid #d9d9d9; /* Cinza claro */
                        border-collapse: collapse;
                    }

                    /* Garante espaçamento e formatação das células */
                    .tabela {
                        border-spacing: 0;
                        border-collapse: collapse;
                    }

                    /* Formata os valores numéricos para terem duas casas decimais */
                    .tabela td {
                        font-variant-numeric: tabular-nums;
                    }
                </style>
            """

            # Criar o email no Outlook
            try:
                outlook = win32.Dispatch("Outlook.Application")
                email = outlook.CreateItem(0)
                email.To = "dados13@grupounus.com.br; supervisao.dados@grupounus.com.br; dados03@grupounus.com.br"
                email.Subject = f'[ACOMPANHAMENTO DIÁRIO DE PLACAS] - Relatório de placas ativadas do dia {yesterday_format}'
                email.HTMLBody = f"""
                    <html>
                    <head>
                        {css_style}  
                    </head>
                    <body>
                        <p>Prezado(a),</p>
                        <p>A seguir, o montante de placas ativas por empresa do dia {yesterday_format}, bem como suas movimentações:</p>

                        {tabela_html}  

                        <p>Atenciosamente,</p>
                        <p>Equipe Análise de Dados - Grupo Unus</p>
                        <p><i>Este é um e-mail automático, por favor não responda</i></p>
                    </body>
                    </html>
                """

                email.Send()
                print("Email enviado com sucesso!")
                logging.info('\n ----------------------------------------------------------------------------------')
                logging.info('\n Processo de Carregamento de Dados 3 concluido com sucesso!')

            except Exception as e:
                print(f"Erro ao enviar o e-mail: {e}")

    def to_whatsapp(df_ativ):   

        # Pega a data mais recente da coluna61981109691
        pyautogui.hotkey('win', 'e')
        time.sleep(2)
        pyautogui.hotkey('shift', 'tab')
        time.sleep(2)
        for _ in range(3):  
            pyautogui.press('down')
        pyautogui.press('enter')
        time.sleep(1)
        # Pasta documents
        pyautogui.hotkey('ctrl', 'e')
        pyautogui.hotkey('shift', 'tab')
        pyautogui.press('right')
        pyautogui.hotkey('shift', 'down')
        for _ in range(7):
            pyautogui.press('down')
        pyautogui.press('enter')
        time.sleep(1)
        # Pasta processos
        pyautogui.hotkey('ctrl', 'e')
        pyautogui.hotkey('shift', 'tab')
        for _ in range(2):
            pyautogui.press('right')                
        for _ in range(6):
            pyautogui.press('down') 
        pyautogui.press('enter')
        time.sleep(1)
        # Pasta envio_wpp
        pyautogui.hotkey('ctrl', 'e')
        time.sleep(1)
        pyautogui.write('tabela_placas_ativas.png')
        time.sleep(2)
        for _ in range(4):
            pyautogui.press('tab')
        time.sleep(1)
        pyautogui.press('down')
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(1)
        #whatsapp
        pyautogui.press('win')
        time.sleep(1)
        pyautogui.write('whatsapp')
        time.sleep(1)
        pyautogui.press('enter')  
        time.sleep(4)
        pyautogui.write('61981109691')  
        time.sleep(1)
        pyautogui.press('down')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.write(f'Tabela de placas ativas do dia {yesterday_format}')  
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)
        for _ in range(4):
            pyautogui.press('tab')
        time.sleep(1)
        pyautogui.press('enter')    

        logging.info('\n ----------------------------------------------------------------------------------')
        logging.info('\n Processo de Carregamento de Dados 4 concluido com sucesso!')


if __name__ == '__main__':
    #devo instanciar pq df_ativ é um atributo de instância (ou de método = self.df_ativ), e não de classe
    #instanciar significa chamar a classe primeiro e colocar em uma variável, isso é criar uma instância de classe
    #se fosse chamar com Extract.extract_dataframe() = chama o método diretamente da classe

    extract_instance = Extract() #chamando a instância de classe

    tabela_df = extract_instance.extract_dataframe()  # chama tabela_df instanciada

    #agora eu posso usar a instância como parâmetro direto nos outros arquivos

    Transform.transform_dataframe(tabela_df)  # Passa tabela_df para Transform

    if today.weekday() not in (5, 6):
        Load.to_outlook(tabela_df, extract_instance.df_ativ)  # Passa tabela_df e df_ativ (atributo de instância) para Load
        #Load.to_whatsapp(extract_instance.df_ativ) 
