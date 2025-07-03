from botcity.core import DesktopBot
import time
import pandas as pd
from datetime import datetime
import os
import sys
from botcity.maestro import *

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False

def not_found(label):
    print(f"Elemento não encontrado: {label}")
    raise Exception(f"Elemento não encontrado: {label}")

# def enviar_mensagem_erro(bot, erro):
#     try:
#         mensagem = f"⚠️ *ERRO NA EXECUÇÃO DO BOT* ⚠️\n\n{erro}"
#         enviar_whatsapp(bot, mensagem)
#     except Exception as e:
#         print(f"Falha ao enviar mensagem de erro: {e}")

def get_excel_path():
    """Retorna o caminho absoluto para o arquivo Excel"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(script_dir, "ContatosGEST.xlsx")

def load_contacts():
    """Carrega a tabela de contatos com tratamento de erros"""
    excel_path = get_excel_path()

    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Arquivo não encontrado em {excel_path}")

    try:
        return pd.read_excel(excel_path)
    except Exception as e:
        raise Exception(f"ERRO ao ler arquivo Excel: {e}")



def login_portal(bot):
    try:
        bot.browse("https://portalanalysisbi.com/login")
        time.sleep(3)
        bot.paste("lucas.alves@adtsa.com.br")
        bot.tab()
        bot.paste("Adtsa2025@@")
        bot.enter()
        time.sleep(5)

        # Searching for element 'Novos '
        if not bot.find("Novos", matching=0.97, waiting_time=10000):
            not_found("Novos")
        bot.click()
                
    except Exception as e:
        raise Exception(f"Falha no login: {str(e)}")

def obter_planilha(bot, nome_arquivo):
    try:

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
            raise FileNotFoundError(f"Arquivo não encontrado em {file_path}")

        aba4 = pd.read_excel(file_path, sheet_name='VOLUME FATURADO NO DIA_3')
        aba5 = pd.read_excel(file_path, sheet_name='VOLUME FATURADO NO DIA_4')
        aba6 = pd.read_excel(file_path, sheet_name='VOLUME FATURADO NO DIA_5')
        aba7 = pd.read_excel(file_path, sheet_name='VOLUME FATURADO NO DIA_6')

        def get_value_or_dash(df, row, col):
            try:
                if df.empty or len(df.columns) <= col or len(df) <= row:
                    return "-"
                value = df.iloc[row, col]
                return value if not pd.isna(value) else "-"
            except:
                return "-"

        total_diario = get_value_or_dash(aba6, 0, 0)
        total_mensal = get_value_or_dash(aba7, 0, 0)

        tipos = ['Direta', 'Novo', 'Usado']
        vendas_diarias = {tipo: get_value_or_dash(aba4, i, 1) for i, tipo in enumerate(tipos) if i < len(aba4)}
        vendas_mensais = {tipo: get_value_or_dash(aba5, i, 1) for i, tipo in enumerate(tipos) if i < len(aba5)}

        data_hoje = datetime.now().strftime('%d/%m/%Y')
        data_mensal = datetime.now().strftime('%m/%Y')

        informe_base = f"""
*INFORMATIVO DE VENDAS VEÍCULOS*

Venda diária - {data_hoje}:
● VDI: {vendas_diarias.get('Direta', '-')}
● VN: {vendas_diarias.get('Novo', '-')}
● VU: {vendas_diarias.get('Usado', '-')}
● Total: {total_diario}

Vendas Mensal - {data_mensal}:
● VDI: {vendas_mensais.get('Direta', '-')}
● VN: {vendas_mensais.get('Novo', '-')}
● VU: {vendas_mensais.get('Usado', '-')}
● Total: {total_mensal}
"""
        return informe_base

    except Exception as e:
        raise Exception(f"Erro ao obter planilha {nome_arquivo}: {str(e)}")


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

    file_path = fr'C:\Users\adtsa\Downloads\{nome_arquivo}.xlsx'

    if not os.path.exists(file_path):
        print(f"ERRO: Arquivo não encontrado em {file_path}")
        return "⚠️ Arquivo não encontrado"

    try:
        aba_volume = pd.read_excel(file_path, sheet_name='VOLUME FATURADO NO DIA', engine='openpyxl')
        aba_chart = pd.read_excel(file_path, sheet_name='Chart 1', engine='openpyxl', header=None)

        print("\nDEBUG - Conteúdo das abas:")
        print("Aba VOLUME FATURADO NO DIA:")
        print(aba_volume)
        print("\nAba Chart 1:")
        print(aba_chart)

        # Função para extrair valores
        def get_value(df, row, col, default="-"):
            try:
                if df.empty or len(df.columns) <= col or len(df) <= row:
                    return default
                value = df.iloc[row, col]
                return default if pd.isna(value) else value
            except:
                return default

        # Extrair totais - CORREÇÃO AQUI: garantindo que pega as colunas corretas
        total_pendentes = get_value(aba_volume, 0, 0)  # Primeira coluna
        valor_total = get_value(aba_volume, 0, 1)      # Segunda coluna

        # Formatar valor monetário - CORREÇÃO: verifica se é numérico antes de formatar
        try:
            valor_num = float(valor_total)
            valor_formatado = f"R$ {valor_num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except (ValueError, TypeError):
            valor_formatado = "-"

        # Processar por tipo de veículo
        if tipo_veiculo == 'caminhoes':
            # Para caminhões - apenas OMODA/JAECOO e Volks
            omoda_jaecoo = "-"
            volks = "-"

            if not aba_chart.empty and len(aba_chart) >= 2:
                for col in range(len(aba_chart.columns)):
                    cabecalho = str(get_value(aba_chart, 0, col, "")).upper()
                    valor = get_value(aba_chart, 1, col)

                    if "OMODA" in cabecalho or "JAECOO" in cabecalho:
                        omoda_jaecoo = valor
                    elif "VOLKS" in cabecalho:
                        volks = valor

            informe = f"""
Total de pedidos pendentes:
● OMODA | JAECOO: {omoda_jaecoo}
● Volks: {volks}
● Total: {total_pendentes}
● Valor: {valor_formatado}
"""
        else:
            # Para carros - apenas as marcas solicitadas
            marcas_procuradas = {
                'VOLKS': 'Volks',
                'RENAULT': 'Renault',
                'GM': 'GM',
                'FORD': 'Ford',
                'CITROEN': 'Citroen',
                'PEUGEOT': 'Peugeot'
            }

            # Inicializa todas as marcas com "-"
            valores = {v: "-" for v in marcas_procuradas.values()}

            if not aba_chart.empty and len(aba_chart) >= 2:
                for col in range(len(aba_chart.columns)):
                    cabecalho = str(get_value(aba_chart, 0, col, "")).upper()
                    valor = get_value(aba_chart, 1, col)

                    for marca_key, marca_nome in marcas_procuradas.items():
                        if marca_key in cabecalho:
                            valores[marca_nome] = valor
                            break

            informe = f"""
Total de pedidos pendentes:
● Volks: {valores['Volks']}
● Renault: {valores['Renault']}
● GM: {valores['GM']}
● Citroen: {valores['Citroen']}
● Peugeot: {valores['Peugeot']}
● Ford: {valores['Ford']}
● Total: {total_pendentes}
● Valor: {valor_formatado}
"""
        return informe

    except Exception as e:
        print(f"Erro ao processar pendentes: {e}")
        return "\nTotal de pedidos pendentes:\n⚠️ Dados não disponíveis no momento"


def sair_entrar(bot):
    try:
        # Searching for element 'BtSair '
        if not bot.find("BtSair", matching=0.97, waiting_time=10000):
            not_found("BtSair")
        bot.click()

        # Searching for element 'EntraCaminhões '
        if not bot.find("EntraCaminhões", matching=0.97, waiting_time=10000):
            not_found("EntraCaminhões")
        bot.click()
    except Exception as e:
        raise Exception(f"Falha ao trocar de portal: {str(e)}")

def enviar_whatsapp(bot, mensagem):
    # Carrega os contatos da planilha
    tabela = load_contacts()
    if tabela is None:
        print("Não foi possível carregar os contatos. Encerrando execução.")
        return

    print("Contatos carregados com sucesso:")
    print(tabela)

    # Abre o WhatsApp Web
    bot.browse("https://web.whatsapp.com/")
    time.sleep(15)  # Tempo para carregar o WhatsApp Web

    # Para cada contato na planilha
    for linha in tabela.index:
        contato = tabela.loc[linha, "contato"]

        try:
            print(f"Processando contato: {contato}")

            # Searching for element 'Lupinha '
            if not bot.find("Lupinha", matching=0.97, waiting_time=10000):
                not_found("Lupinha")
            bot.click()


            # Limpa o campo de pesquisa e digita o contato
            bot.control_a()
            bot.backspace()
            bot.type_keys_with_interval(100,    str(contato))
            time.sleep(3)  # Espera os resultados aparecerem

            bot.enter()
            time.sleep(5)  # Tempo para carregar a conversa

            # Envia a mensagem unificada
            bot.paste(mensagem)
            time.sleep(1)
            bot.enter()
            time.sleep(3)  # Espera a mensagem ser enviada

            print(f"Mensagem enviada para: {contato}")

            # Espera um pouco antes do próximo contato
            time.sleep(3)

        except Exception as e:
            print(f"Erro ao enviar para {contato}: {str(e)}")
            # Tenta continuar para o próximo contato
            continue


def main():
    bot = DesktopBot()
    erro_global = None

    try:
        # ========== CARROS ==========
        login_portal(bot)
        informe_vendas_carros = obter_planilha(bot, "VN Carros")
        informe_pendentes_carros = obter_planilha_pendentes(bot, "VN Carros Pendentes", 'carros')

        # ========== CAMINHÕES ==========
        sair_entrar(bot)
        time.sleep(5)
        informe_vendas_caminhoes = obter_planilha(bot, "VN Caminhões")
        informe_pendentes_caminhoes = obter_planilha_pendentes(bot, "VN Caminhões pendentes", 'caminhoes')

        # ========== MENSAGEM FINAL ==========
        mensagem_unificada = (
            f"{informe_vendas_carros}\n"
            f"{informe_pendentes_carros}\n"
            f"*CAMINHÕES*\n"
            f"{informe_vendas_caminhoes}\n"
            f"{informe_pendentes_caminhoes}"
        )

        enviar_whatsapp(bot, mensagem_unificada)

    except Exception as e:
        # erro_global = str(e)
        # print(f"ERRO CRÍTICO: {erro_global}")
        #
        # # Tenta notificar o erro mesmo que tenha falhado em outras etapas
        # try:
        #     if "Falha ao enviar mensagem no WhatsApp" not in erro_global:
        #         #enviar_mensagem_erro(bot, erro_global)
        # except Exception as e2:
        #     print(f"Falha ao tentar notificar erro: {e2}")

        # Encerra com erro
        sys.exit(1)


if __name__ == "__main__":
    main()








