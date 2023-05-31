import requests
from openpyxl import load_workbook, Workbook

# Função para pesquisar e capturar os dados do CNPJ
def pesquisar_cnpj(cnpj):
    url = f"https://www.receitaws.com.br/v1/cnpj/{cnpj}"
    response = requests.get(url)
    data = response.json()

    # Capturar os dados relevantes
    nome_empresa = data["nome"]
    nome_fantasia = data["fantasia"]
    endereco = f"{data['logradouro']}, {data['numero']} - {data['bairro']}"
    telefone = data["telefone"]
    email = data["email"]
    area_atuacao = data["atividade_principal"][0]["text"]

    # Retornar os dados capturados
    return [nome_empresa, nome_fantasia, endereco, telefone, email, area_atuacao]

try:
    # Carregar a planilha existente
    planilha = load_workbook("dados_empresas.xlsx")
    print("Planilha existente carregada com sucesso!")
except FileNotFoundError:
    # Caso o arquivo não exista, criar uma nova planilha vazia
    planilha = Workbook()
    planilha.active.append(["Nome da Empresa", "Nome Fantasia", "Endereço", "Telefone", "Email", "Área de Atuação"])
    print("Nova planilha criada!")

sheet = planilha.active

while True:
    # Solicitar CNPJ ao usuário
    cnpj = input("Digite o CNPJ da empresa (ou 'sair' para encerrar): ")

    if cnpj.lower() == "sair":
        break

    # Pesquisar e capturar os dados do CNPJ fornecido
    dados_cnpj = pesquisar_cnpj(cnpj)

    # Adicionar os dados à planilha
    sheet.append(dados_cnpj)

    # Ajustar o tamanho das células
    for column_cells in sheet.columns:
        max_length = 0
        for cell in column_cells:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        sheet.column_dimensions[column_cells[0].column_letter].width = adjusted_width

# Salvar a planilha em um arquivo
planilha.save("dados_empresas.xlsx")
print("Dados salvos na planilha com sucesso!")
