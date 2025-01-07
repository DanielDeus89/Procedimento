import openpyxl
import os
import datetime
import shutil
import requests
from tqdm import tqdm  # Importando tqdm para a barra de progresso

def download_file(url, local_filename):
    with requests.get(url, stream=True) as r:
        r.raise_for_status()
        with open(local_filename, 'wb') as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)
    return local_filename

def criar_copias(identificadores, categoria, meses, mes_atual, template_filename, output_dir, cell_categoria, cell_id, cell_mes):
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)

    # Adicionando a barra de progresso
    for id in tqdm(identificadores, desc=f"Copiando arquivos para {categoria}", unit="arquivo"):
        arquivo = openpyxl.load_workbook(template_filename)
        planilha = arquivo.active

        planilha[cell_categoria] = categoria
        planilha[cell_id] = id
        planilha[cell_mes] = meses[mes_atual - 1]
        arquivo.save(os.path.join(output_dir, id + ".xlsx"))
        arquivo.close()

def processar_template(url, local_filename, identificadores_categoria, output_dir, cell_coords):
    # Fazer o download do arquivo
    download_file(url, local_filename)

    # Limpeza de diretório antigo e criação de novo
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.mkdir(output_dir)

    # Limpeza de arquivos antigos
    for filename in os.listdir():
        if filename.endswith(".xlsx") and (filename.startswith("Smar") or filename.startswith("Adda")):
            os.remove(filename)

    # Renomeação do arquivo principal
    for filename in os.listdir():
        if filename.endswith(".xlsx") and filename != local_filename:
            os.rename(filename, local_filename)
            break  # Assumindo que só existe um arquivo .xlsx relevante

    # Obter mês atual
    meses = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"]
    mes_atual = datetime.date.today().month

    # Criação de cópias com barra de progresso
    for categoria, identificadores in identificadores_categoria.items():
        criar_copias(identificadores, categoria, meses, mes_atual, local_filename, output_dir, *cell_coords)

    # Remoção do arquivo original
    os.remove(local_filename)

def main():
    # Definir os parâmetros para cada template
    templates = [
        {
            "url": "http://10.8.2.100:8094/edoc/DownloadDocumentFile?documentProperties=eyJEb2N1bWVudENvZGUiOiJGQlItRU4tNC00NDgiLCJTaG91bGRJbXBlcnNvbmF0ZSI6dHJ1ZSwiRG9tYWluIjoibGEuZm94Y29ubi5jb20iLCJVc2VyTmFtZSI6InBvcnRhbC1xYS1maWxlcyIsIlBhc3N3b3JkIjoiUDBydEAhNyIsIkZpbGVOYW1lIjoiRkJSLUVOLTQtNDQ4LnBkZiIsIkZ1bGxQYXRoIjoiXFxcXDEwLjguMi4zMFxcRmlsZXMkXFxGQlJcXFF1YWxpdHlcXERvY3VtZW50b3MgZSBSZWdpc3Ryb3MgU0dJXFxET0NVTUVOVE9TIE9GSUNJQUlTXFxfRURPQ1xcUFVCTElDQURPU1xcRkJSTEFcXEZCUi1FTi00LTQ0OFxcODk3MDNcXEZCUi1FTi00LTQ0OF80Lnhsc3giLCJJRCI6MCwiVXNlcklEIjowfQ==",
            "local_filename": "FBR-EN-4-448.xlsx",
            "identificadores_categoria": {
                'FT1-MP1': ['SmarX1125A4769', 'SmarX1125A3432', 'SmarX1125A8970', 'SmarX1125A2599', 'SmarX1125A5724'],
                'FT2-MP1': ['SmarX1188A5001', 'SmarX1188A5002', 'SmarX1188A1026', 'SmarX1188A8006', 'SmarX8700A9179', 'SmarX8700A1023'],
                'FT2-MP2': ['Addax4977A0172', 'SmarX1125A1276', 'SmarX1125A2827', 'SmarX1125A2837', 'SmarX1125A7804'],
                'FT2-MP3': ['SmarX1125A1278', 'SmarX1125A2848', 'SmarX1125A7800', 'SmarX1125A8047'],
                'MGT-MP6': ['SmarX1125A5825', 'SmarX1125A7192'],
                'ST-MP1': ['SmarX1280A4702', 'SmarX1280A0375'],
                'ST-MP3': ['SmarX1288A0067'],
                'BURN IN': ['AddaX8754A0140', 'AddaX8754A0145']
            },
            "output_dir": "FBR-EN-4-448",
            "cell_coords": ("C4", "F4", "AA4")
        },
        {
            "url": "http://10.8.2.100:8094/edoc/DownloadDocumentFile?documentProperties=eyJEb2N1bWVudENvZGUiOiJGQlItRU4tNC00NTAiLCJTaG91bGRJbXBlcnNvbmF0ZSI6dHJ1ZSwiRG9tYWluIjoibGEuZm94Y29ubi5jb20iLCJVc2VyTmFtZSI6InBvcnRhbC1xYS1maWxlcyIsIlBhc3N3b3JkIjoiUDBydEAhNyIsIkZpbGVOYW1lIjoiRkJSLUVOLTQtNDUwLnBkZiIsIkZ1bGxQYXRoIjoiXFxcXDEwLjguMi4zMFxcRmlsZXMkXFxGQlJcXFF1YWxpdHlcXERvY3VtZW50b3MgZSBSZWdpc3Ryb3MgU0dJXFxET0NVTUVOVE9TIE9GSUNJQUlTXFxfRURPQ1xcUFVCTElDQURPU1xcRkJSTEFcXEZCUi1FTi00LTQ1MFxcMTE5Mzc5XFxGQlItRU4tNC00NTBfNC54bHN4IiwiSUQiOjAsIlVzZXJJRCI6MH0=",
            "local_filename": "FBR-EN-4-450.xlsx",
            "identificadores_categoria": {
                'FT1-MP1': ['SmarX1125A4769', 'SmarX1125A3432', 'SmarX1125A8970', 'SmarX1125A2599', 'SmarX1125A5724'],
                'FT2-MP2': ['SmarX1125A1276', 'SmarX1125A2827', 'SmarX1125A2837', 'SmarX1125A7804'],
                'FT2-MP3': ['SmarX1125A1278', 'SmarX1125A2848', 'SmarX1125A7800', 'SmarX1125A8047']
            },
            "output_dir": "FBR-EN-4-450",
            "cell_coords": ("B4", "D4", "G4")
        },
        {
            "url": "http://10.8.2.100:8094/edoc/DownloadDocumentFile?documentProperties=eyJEb2N1bWVudENvZGUiOiJGQlItRU4tNC00ODkiLCJTaG91bGRJbXBlcnNvbmF0ZSI6dHJ1ZSwiRG9tYWluIjoibGEuZm94Y29ubi5jb20iLCJVc2VyTmFtZSI6InBvcnRhbC1xYS1maWxlcyIsIlBhc3N3b3JkIjoiUDBydEAhNyIsIkZpbGVOYW1lIjoiRkJSLUVOLTQtNDg5LnBkZiIsIkZ1bGxQYXRoIjoiXFxcXDEwLjguMi4zMFxcRmlsZXMkXFxGQlJcXFF1YWxpdHlcXERvY3VtZW50b3MgZSBSZWdpc3Ryb3MgU0dJXFxET0NVTUVOVE9TIE9GSUNJQUlTXFxfRURPQ1xcUFVCTElDQURPU1xcRkJSTEFcXEZCUi1FTi00LTQ4OVxcODk3MDZcXEZCUi1FTi00LTQ4OV8zLnhsc3giLCJJRCI6MCwiVXNlcklEIjowfQ==",
            "local_filename": "FBR-EN-4-489.xlsx",
            "identificadores_categoria": {
                'RUNIN': ['RUNIN'],
                'BURN IN': ['AddaX8754A0140', 'AddaX8754A0145']
            },
            "output_dir": "FBR-EN-4-489",
            "cell_coords": ("C4", "F4", "AA4")
        },
        {
            "url": "http://10.8.2.100:8094/edoc/DownloadDocumentFile?documentProperties=eyJEb2N1bWVudENvZGUiOiJGQlItRU4tNC00OTAiLCJTaG91bGRJbXBlcnNvbmF0ZSI6dHJ1ZSwiRG9tYWluIjoibGEuZm94Y29ubi5jb20iLCJVc2VyTmFtZSI6InBvcnRhbC1xYS1maWxlcyIsIlBhc3N3b3JkIjoiUDBydEAhNyIsIkZpbGVOYW1lIjoiRkJSLUVOLTQtNDkwLnBkZiIsIkZ1bGxQYXRoIjoiXFxcXDEwLjguMi4zMFxcRmlsZXMkXFxGQlJcXFF1YWxpdHlcXERvY3VtZW50b3MgZSBSZWdpc3Ryb3MgU0dJXFxET0NVTUVOVE9TIE9GSUNJQUlTXFxfRURPQ1xcUFVCTElDQURPU1xcRkJSTEFcXEZCUi1FTi00LTQ5MFxcODk3MDdcXEZCUi1FTi00LTQ5MF8zLnhsc3giLCJJRCI6MCwiVXNlcklEIjowfQ==",
            "local_filename": "FBR-EN-4-490.xlsx",
            "identificadores_categoria": {
                'FT2-MP1': ['SmarX8700A2018', 'SmarX8700A8116']
            },
            "output_dir": "FBR-EN-4-490",
            "cell_coords": ("C4", "J4", "AA4")
        }
    ]

    # Processar cada template
    for template in templates:
        processar_template(template['url'], template['local_filename'], template['identificadores_categoria'], template['output_dir'], template["cell_coords"])

if __name__ == "__main__":
    main()
