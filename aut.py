from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

arquivo = "vendas.xlsx"

try:
    planilha = load_workbook(arquivo)
    pagina = planilha.active
    print("Planilha carregada!")

except:
    planilha = Workbook()
    pagina = planilha.active
    pagina.title = "Vendas"

    pagina["A1"] = "Produto"
    pagina["B1"] = "Quantidade"
    pagina["C1"] = "Preço"
    pagina["D1"] = "Total"

    for celula in pagina[1]:
        celula.font = Font(bold=True, color="FFFFFF")
        celula.fill = PatternFill(
            start_color="0000FF",
            end_color="0000FF",
            fill_type="solid"
        )

    print("Nova planilha criada!")

produto = input("Produto: ")
quantidade = int(input("Quantidade: "))
preco = float(input("Preço: "))

total = quantidade * preco

pagina.append([produto, quantidade, preco, total])

planilha.save(arquivo)

print("Dados salvos com sucesso!")