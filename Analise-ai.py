import os
import pandas as pd
import PyPDF2
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

class PedidoHandler(FileSystemEventHandler):
    def __init__(self, arquivo_excel):
        self.arquivo_excel = arquivo_excel

    def on_created(self, event):
        if event.is_directory:
            return
        
        if event.src_path.endswith(".pdf"):
            arquivo_pdf = event.src_path
            print(f"Novo pedido de venda encontrado: {arquivo_pdf}")
            arquivo_saida = "novo_pedido.xlsx"  # Nome do arquivo de saída
            processar_pedido(self.arquivo_excel, arquivo_pdf, arquivo_saida)

def ler_excel(arquivo_excel):
    df = pd.read_excel(arquivo_excel)
    return df

def ler_pdf(arquivo_pdf):
    with open(arquivo_pdf, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        texto = ""
        for page in reader.pages:
            texto += page.extract_text()
    return texto.splitlines()

def gerar_novo_excel(df_pecas, pedido_venda, arquivo_saida):
    dados_saida = []
    itens_nao_encontrados = []

    for linha in pedido_venda:
        codigo_peca = linha.strip()  # Ajuste conforme necessário
        componentes = df_pecas[df_pecas['Nome da Peça'] == codigo_peca]

        if not componentes.empty:
            for _, componente in componentes.iterrows():
                dados_saida.append({
                    "Código": componente['Código'],
                    "Quantidade": componente['Quantidade'],
                    "Unidade": componente['Unidade'],
                    "Descrição": componente['Descrição'],
                    "Nome da Peça": componente['Nome da Peça']
                })
        else:
            itens_nao_encontrados.append(codigo_peca)  # Registra item não encontrado

    df_saida = pd.DataFrame(dados_saida)
    df_saida.to_excel(arquivo_saida, index=False)

    # Exibir itens não encontrados, se houver
    if itens_nao_encontrados:
        print("Os seguintes itens não foram encontrados no banco de dados:")
        for item in itens_nao_encontrados:
            print(item)

def processar_pedido(arquivo_excel, arquivo_pdf, arquivo_saida):
    df_pecas = ler_excel(arquivo_excel)
    pedido_venda = ler_pdf(arquivo_pdf)
    gerar_novo_excel(df_pecas, pedido_venda, arquivo_saida)

def monitorar_pasta(pasta_monitorada, arquivo_excel):
    event_handler = PedidoHandler(arquivo_excel)
    observer = Observer()
    observer.schedule(event_handler, pasta_monitorada, recursive=False)
    observer.start()
    print(f"Monitorando a pasta: {pasta_monitorada}")

    try:
        while True:
            pass  # Mantém o script em execução
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

# Configurações
pasta_monitorada = "caminho/para/pasta/pedido_venda"  # Altere para o caminho da pasta que deseja monitorar
arquivo_excel = "banco_dados.xlsx"  # Banco de dados fixo

monitorar_pasta(pasta_monitorada, arquivo_excel)
