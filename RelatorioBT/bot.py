from botcity.core import DesktopBot
import time
import pandas as pd
from datetime import datetime
import os
from botcity.maestro import *

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False

def not_found(label):
    print(f"Element not found: {label}")
    #bot.stop_browser()
    exit()

def login_portal(bot):
    bot.browse("https://portalanalysisbi.com/login")
    time.sleep(3)
    bot.paste("lucas.alves@adtsa.com.br")
    bot.tab()
    bot.paste("Adtsa2025@@")
    bot.enter()
    time.sleep(5)


    # Searching for element 'EntrarNOV '
    if not bot.find("EntrarNOV", matching=0.97, waiting_time=10000):
        not_found("EntrarNOV")
    bot.click()



def obter_planilha(bot, nome_arquivo):


    # Searching for element 'VendasMove '
    if not bot.find("VendasMove", matching=0.97, waiting_time=30000):
        not_found("VendasMove")
    bot.move()

    # Searching for element 'BtExportar '
    if not bot.find("BtExportar", matching=0.97, waiting_time=10000):
        not_found("BtExportar")
    bot.click()

    # Searching for element 'ExportExcel '
    if not bot.find("ExportExcel", matching=0.97, waiting_time=10000):
        not_found("ExportExcel")
    bot.click()

    # Searching for element 'RelativoNome '
    if not bot.find("RelativoNome", matching=0.97, waiting_time=10000):
        not_found("RelativoNome")
    bot.click_relative(239, 13)

    bot.type_keys(["ctrl", "a"])
    bot.backspace()
    bot.paste(nome_arquivo)


    # Searching for element 'BaixarPlanilha '
    if not bot.find("BaixarPlanilha", matching=0.97, waiting_time=10000):
        not_found("BaixarPlanilha")
    bot.click()

    time.sleep(10)

    # Processamento do arquivo
    file_path = fr'C:\Users\adtsa\Downloads\{nome_arquivo}.xlsx'

    if not os.path.exists(file_path):
        print(f"ERRO: Arquivo não encontrado em {file_path}")
        print("Por favor, verifique:")
        print(f"1. O arquivo '{nome_arquivo}.xlsx' está na pasta 'Downloads'?")
        print("2. O nome do arquivo está correto?")
        print("3. O arquivo não está aberto em outro programa?")
        exit()

    try:
        aba4 = pd.read_excel(file_path, sheet_name='VOLUME FATURADO NO DIA_3')
        aba5 = pd.read_excel(file_path, sheet_name='VOLUME FATURADO NO DIA_4')
        aba6 = pd.read_excel(file_path, sheet_name='VOLUME FATURADO NO DIA_5')
        aba7 = pd.read_excel(file_path, sheet_name='VOLUME FATURADO NO DIA_6')

        vendas_diarias = aba4.set_index('Estoque_Tipo')['Qtde (Soma)'].to_dict()
        vendas_mensais = aba5.set_index('Estoque_Tipo')['Filtered Qtde (Soma)'].to_dict()

        data_hoje = datetime.now().strftime('%d/%m/%Y')
        data_mensal = datetime.now().strftime('%m/%Y')

        informe_base = f"""
Venda diária - {data_hoje}:
● VDI: {vendas_diarias.get('Direta', 0)}
● VN: {vendas_diarias.get('Novo', 0)}
● VU: {vendas_diarias.get('Usado', 0)}
● Total: {aba6.iloc[0, 0]}

Vendas Mensal - {data_mensal}:
● VDI: {vendas_mensais.get('Direta', 0)}
● VN: {vendas_mensais.get('Novo', 0)}
● VU: {vendas_mensais.get('Usado', 0)}
● Total: {aba7.iloc[0, 0]}
    """
        return informe_base

    except Exception as e:
        print(f"Erro ao processar o arquivo: {e}")
        print("Possíveis causas:")
        print("- Nomes das abas estão diferentes do esperado")
        print("- Estrutura do arquivo foi alterada")
        print("- Arquivo corrompido")
        exit()


def obter_planilha_pendentes(bot, nome_arquivo, tipo_veiculo='carros'):

    # Searching for element 'PendenteMoves '
    if not bot.find("PendenteMoves", matching=0.97, waiting_time=10000):
        not_found("PendenteMoves")
    bot.click()

    # Searching for element 'BtExportar '
    if not bot.find("BtExportar", matching=0.97, waiting_time=10000):
        not_found("BtExportar")
    bot.click()

    # Searching for element 'ExportExcel '
    if not bot.find("ExportExcel", matching=0.97, waiting_time=10000):
        not_found("ExportExcel")
    bot.click()

    # Searching for element 'RelativoNome '
    if not bot.find("RelativoNome", matching=0.97, waiting_time=10000):
        not_found("RelativoNome")
    bot.click_relative(239, 13)

    bot.type_keys(["ctrl", "a"])
    bot.backspace()
    bot.paste(nome_arquivo)


    # Searching for element 'BaixarPlanilha '
    if not bot.find("BaixarPlanilha", matching=0.97, waiting_time=10000):
        not_found("BaixarPlanilha")
    bot.click()

    time.sleep(10)

    # Processamento do arquivo
    file_path = fr'C:\Users\adtsa\Downloads\{nome_arquivo}.xlsx'

    if not os.path.exists(file_path):
        print(f"ERRO: Arquivo não encontrado em {file_path}")
        print("Por favor, verifique:")
        print(f"1. O arquivo '{nome_arquivo}.xlsx' está na pasta 'Downloads'?")
        print("2. O nome do arquivo está correto?")
        print("3. O arquivo não está aberto em outro programa?")
        exit()

    try:
        aba_volume = pd.read_excel(file_path, sheet_name='VOLUME FATURADO NO DIA')
        aba_chart = pd.read_excel(file_path, sheet_name='Chart 1')

        # Extrai dados básicos
        total_pendentes = aba_volume.iloc[0, 0] if len(aba_volume.columns) > 0 else 0
        valor_total = aba_volume.iloc[0, 1] if len(aba_volume.columns) > 1 else 0
        valor_formatado = f"R$ {valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        # Processa as marcas de acordo com o tipo de veículo
        if tipo_veiculo == 'caminhoes':
            # Para caminhões: apenas Volks e OMODA | JAECOO
            volks = 0
            omoda_jaecoo = 0

            for col in aba_chart.columns:
                col_name = str(col).strip()
                if 'volks' in col_name.lower():
                    volks = aba_chart[col].iloc[0] if not pd.isna(aba_chart[col].iloc[0]) else 0
                elif 'omoda' in col_name.lower() or 'jaecoo' in col_name.lower():
                    omoda_jaecoo = aba_chart[col].iloc[0] if not pd.isna(aba_chart[col].iloc[0]) else 0

            informe = f"""
Total de pedidos pendentes:
● Volks: {volks}
● OMODA | JAECOO: {omoda_jaecoo}
● Total: {total_pendentes}
● Valor: {valor_formatado}
"""
        else:
            # Para carros: todas as marcas
            marcas = {
                'Volks': 0,
                'GM': 0,
                'Ford': 0,
                'Renault': 0,
                'Citroen': 0,
                'Peugeot': 0
            }

            for col in aba_chart.columns:
                col_name = str(col).strip().lower()
                if 'volks' in col_name:
                    marcas['Volks'] = aba_chart[col].iloc[0] if not pd.isna(aba_chart[col].iloc[0]) else 0
                elif 'gm' in col_name or 'chevrolet' in col_name:
                    marcas['GM'] = aba_chart[col].iloc[0] if not pd.isna(aba_chart[col].iloc[0]) else 0
                elif 'ford' in col_name:
                    marcas['Ford'] = aba_chart[col].iloc[0] if not pd.isna(aba_chart[col].iloc[0]) else 0
                elif 'renault' in col_name:
                    marcas['Renault'] = aba_chart[col].iloc[0] if not pd.isna(aba_chart[col].iloc[0]) else 0
                elif 'citroen' in col_name:
                    marcas['Citroen'] = aba_chart[col].iloc[0] if not pd.isna(aba_chart[col].iloc[0]) else 0
                elif 'peugeot' in col_name:
                    marcas['Peugeot'] = aba_chart[col].iloc[0] if not pd.isna(aba_chart[col].iloc[0]) else 0

            informe = f"""
Total de pedidos pendentes:
● Volks: {marcas['Volks']}
● Renault: {marcas['Renault']}
● GM: {marcas['GM']}
● Citroen: {marcas['Citroen']}
● Peugeot: {marcas['Peugeot']}
● Ford: {marcas['Ford']}
● Total: {total_pendentes}
● Valor: {valor_formatado}

"""
        return informe

    except Exception as e:
        print(f"Erro ao processar o arquivo de pendentes: {str(e)}")
        print("Verifique:")
        print("- Se as abas estão com os nomes corretos")
        print("- Se o formato do arquivo está correto")
        print(f"- Caminho do arquivo: {file_path}")
        exit()


def sair_entrar(bot):

    # Searching for element 'BtSair '
    if not bot.find("BtSair", matching=0.97, waiting_time=10000):
        not_found("BtSair")
    bot.click()
    
    # Searching for element 'EntraCaminhões '
    if not bot.find("EntraCaminhões", matching=0.97, waiting_time=10000):
        not_found("EntraCaminhões")
    bot.click()
    



def enviar_whatsapp(bot, contato, mensagem):

    bot.browse("https://web.whatsapp.com/")

    # Searching for element 'Lupinha '
    if not bot.find("Lupinha", matching=0.97, waiting_time=10000):
        not_found("Lupinha")
    bot.click()
    
    bot.paste(contato)
    time.sleep(2)
    bot.enter()
    time.sleep(2)
    bot.paste(mensagem)
    bot.enter()



def main():
    bot = DesktopBot()

    try:
        # Login no portal
        login_portal(bot)

        # ========== CARROS ==========
        # Obter relatórios de carros
        informe_vendas_carros = obter_planilha(bot, "VN Carros")
        informe_pendentes_carros = obter_planilha_pendentes(bot, "VN Carros Pendentes", 'carros')
        # titulo_carros = "\n*RELATÓRIO DE CARROS*\n"

        # ========== TROCAR PARA CAMINHÕES ==========
        sair_entrar(bot)
        time.sleep(5)  # Esperar a transição

        # ========== CAMINHÕES ==========
        # Obter relatórios de caminhões
        informe_vendas_caminhoes = obter_planilha(bot, "VN Caminhões")
        informe_pendentes_caminhoes = obter_planilha_pendentes(bot, "VN Caminhões pendentes", 'caminhoes')
        # titulo_caminhoes = "\n*RELATÓRIO DE CAMINHÕES*\n"

        # ========== MENSAGEM ÚNICA ==========
        mensagem_unificada = (
            "*INFORMATIVO DE VENDAS VEÍCULOS*"
            #f"{titulo_carros}"
            f"{informe_vendas_carros}"
            f"{informe_pendentes_carros}"
            "*CAMINHÕES*"
            #f"{titulo_caminhoes}"
            f"{informe_vendas_caminhoes}"
            f"{informe_pendentes_caminhoes}"
        )

        # Enviar por WhatsApp em uma única mensagem
        enviar_whatsapp(bot, "Lucas", mensagem_unificada)
        time.sleep(5)

    except Exception as e:
        print(f"Erro durante a execução: {e}")
        bot.stop_browser()
        exit()

if __name__ == "__main__":
    main()



