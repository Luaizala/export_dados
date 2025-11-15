import pandas as pd
import os

# define o nome do arquivo final
arquivo = "registro_itens.xlsx"
produtos = []

def formatar_moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

while True:
    print("\n==== MENU ====")
    print("1 - Adicionar novo produto")
    print("2 - Sair e gerar Excel")

    opcao = input("Escolha uma opção: ")

    if opcao == "1":
        nome_produto = input("Digite o nome do produto: ")

        # Valida quantidade
        while True:
            try:
                quantidade = input("Digite a quantidade comprada: ")
                quantidade_validada = int(quantidade)
                break
            except ValueError:
                print("Por favor, insira um número válido para a quantidade.")

        # Valida preço
        while True:
            try:
                preco = input("Digite o preço unitário do produto: ")
                preco_validado = float(preco)
                break
            except ValueError:
                print("Por favor, insira um número válido para o preço.")

        valor_total = quantidade_validada * preco_validado

        print("\nProduto cadastrado:")
        print("Nome:", nome_produto)
        print("Quantidade:", quantidade_validada)
        print("Preço unitário:", formatar_moeda(preco_validado))
        print("Valor total:", formatar_moeda(valor_total))

        produtos.append({
            "Item": nome_produto,
            "Quantidade": quantidade_validada,
            "Preço Unitário": formatar_moeda(preco_validado),
            "Valor Total": formatar_moeda(valor_total)
        })

    elif opcao == "2":
        print("\nGerando o arquivo Excel...")
        break

    else:
        print("Opção inválida. Tente novamente.")

# === GERAÇÃO DO EXCEL ===

df_novos = pd.DataFrame(produtos)

if not os.path.exists(arquivo):
    df_novos.to_excel(arquivo, index=False)
else:
    df_existente = pd.read_excel(arquivo)
    df_atualizado = pd.concat([df_existente, df_novos], ignore_index=True)
    df_atualizado.to_excel(arquivo, index=False)

print("Arquivo 'registro_itens.xlsx' atualizado com sucesso!")
