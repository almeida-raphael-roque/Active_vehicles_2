import matplotlib.pyplot as plt
from pandas.plotting import table
import pythoncom
import pandas as pd
from extract import Extract
import logging

logging.basicConfig(
    level=logging.INFO,  # Exibe mensagens a partir de INFO
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.StreamHandler()  # Garante logs no console
    ]
)


class Transform:

    def transform_dataframe(tabela_df):

        # Cria a figura para o gráfico
        fig, ax = plt.subplots(figsize=(8, 4))  # Ajuste o tamanho conforme necessário

        # Remove os eixos
        ax.axis('off')

        # Cria a tabela no gráfico
        tbl = table(ax, tabela_df, loc='center', colWidths=[0.2]*len(tabela_df.columns))

        # Ajusta a fonte e o estilo da tabela
        tbl.auto_set_font_size(False)
        tbl.set_fontsize(12)
        tbl.scale(1.2, 1.5)  # Aumenta ligeiramente o espaçamento vertical

        # Formatação de células (estilo CSS implementado no Python)
        for key, cell in tbl.get_celld().items():
            cell.set_text_props(horizontalalignment='center', verticalalignment='center')  # Centraliza tudo
            
            if key[0] == 0:  # Cabeçalhos (primeira linha da tabela)
                cell.set_text_props(fontweight='bold', color='black')  # Ajusta o peso da fonte e cor
                cell.set_facecolor('#d9d9d9')  # Cor de fundo cinza para cabeçalho
            elif key[1] == len(tabela_df.columns) - 1:  # Última coluna
                cell.set_text_props(fontweight='bold', color='black')  # Ajusta o peso da fonte e cor
                cell.set_facecolor('#d9d9d9')  # Cor de fundo cinza para última coluna
            elif key[0] == 6:  # Última linha (índice 6, contando com cabeçalho)
                cell.set_text_props(fontweight='bold', color='black')  # Negrito para a sétima linha
                cell.set_facecolor('#d9d9d9')  # Cor de fundo cinza para a sétima linha
            if key == (6, len(tabela_df.columns) - 1):  # Última célula específica
                cell.set_text_props(fontweight='bold', color='white')  # Texto branco em negrito
                cell.set_facecolor('#4d4d4d')  # Fundo cinza escuro

        # Salva a imagem PNG

        output_path = "tabela_placas_ativas.png"
        plt.savefig(output_path, format='png', bbox_inches='tight', dpi=300)

        
        # Fecha o gráfico
        plt.close()

        print(f"Tabela salva como {output_path}")

# if __name__ ==  '__main__':
#     Transform.transform_dataframe()