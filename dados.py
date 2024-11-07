import pandas as pd
from tabulate import tabulate

def abrir_arquivo():
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', 200)
    df = pd.read_excel('cbc information.xlsx')
    return df

def exibir_informcoes_exames(dado, informacoes_exames):
    if dado in informacoes_exames:
        info = informacoes_exames[dado]
        print(f"Nome: {info['nome']}")
        print(f"Faixa Normal: {info['faixa_normal']}")
        print(f"Unidade: {info['unidade']}")
    else:
        print("Exame não encontrado")

def buscar_exames_por_faixa(df, informacoes_exames, condicao):

    resultados = []
    for coluna in informacoes_exames:
        if informacoes_exames[coluna]["faixa_normal"] is not None:
            faixa = informacoes_exames[coluna]["faixa_normal"].split(" a ")
            min_val = float(faixa[0])
            max_val = float(faixa[1])
            if condicao == "acima":
                resultados.append(df[df[coluna] > max_val][[coluna]])
            elif condicao == "abaixo":
                resultados.append(df[df[coluna] < min_val][[coluna]])
            elif condicao == "nos conformes":
                resultados.append(df[(df[coluna] >= min_val) & (df[coluna] <= max_val)][[coluna]])

    resultados_df = pd.concat(resultados, axis=1).fillna('--')
    return resultados_df

def listar_todos_exames(informacoes_exames):
    print("\nListagem de todos os exames:\n")
    for exame, detalhes in informacoes_exames.items():
        print(f"Exame: {exame}")
        print(f"Nome: {detalhes['nome']}")
        print(f"Faixa Normal: {detalhes['faixa_normal']}")
        print(f"Unidade: {detalhes['unidade']}")
        print("-" * 30)

def listar_tabela(df, quantidade_tabela):
    print(tabulate(df.head(quantidade_tabela), headers='keys', tablefmt='fancy_grid'))

def main():
    df = abrir_arquivo()

    informacoes_exames = {
        "ID": {"nome": "Identificação dos Pacientes", "faixa_normal": None, "unidade": None},
        "WBC": {"nome": "Contagem de Glóbulos Brancos", "faixa_normal": "4.0 a 10.0", "unidade": "10^9/L"},
        "LYMp": {"nome": "Percentual de Linfócitos", "faixa_normal": "20.0 a 40.0", "unidade": "%"},
        "MIDp": {"nome": "Percentual de Outros Glóbulos Brancos", "faixa_normal": "1.0 a 15.0", "unidade": "%"},
        "NEUTp": {"nome": "Percentual de Neutrófilos", "faixa_normal": "50.0 a 70.0", "unidade": "%"},
        "LYMn": {"nome": "Número de Linfócitos", "faixa_normal": "0.6 a 4.1", "unidade": "10^9/L"},
        "MIDn": {"nome": "Número de Outros Glóbulos Brancos", "faixa_normal": "0.1 a 1.8", "unidade": "10^9/L"},
        "NEUTn": {"nome": "Número de Neutrófilos", "faixa_normal": "2.0 a 7.8", "unidade": "10^9/L"},
        "RBC": {"nome": "Contagem de Glóbulos Vermelhos", "faixa_normal": "3.50 a 5.50", "unidade": "10^12/L"},
        "HGB": {"nome": "Hemoglobina", "faixa_normal": "11.0 a 16.0", "unidade": "g/dL"},
        "HCT": {"nome": "Hematócrito", "faixa_normal": "36.0 a 48.0", "unidade": "%"},
        "MCV": {"nome": "Volume Corpuscular Médio", "faixa_normal": "80.0 a 99.0", "unidade": "fL"},
        "MCH": {"nome": "Hemoglobina Corpuscular Média", "faixa_normal": "26.0 a 32.0", "unidade": "pg"},
        "MCHC": {"nome": "Concentração de Hemoglobina Corpuscular Média", "faixa_normal": "32.0 a 36.0", "unidade": "g/dL"},
        "RDWSD": {"nome": "Largura de Distribuição dos Glóbulos Vermelhos", "faixa_normal": "37.0 a 54.0", "unidade": "fL"},
        "RDWCV": {"nome": "Largura de Distribuição dos Glóbulos Vermelhos", "faixa_normal": "11.5 a 14.5", "unidade": "%"},
        "PLT": {"nome": "Contagem de Plaquetas", "faixa_normal": "100 a 400", "unidade": "10^9/L"},
        "MPV": {"nome": "Volume Plaquetário Médio", "faixa_normal": "7.4 a 10.4", "unidade": "fL"},
        "PDW": {"nome": "Largura de Distribuição dos Glóbulos Vermelhos", "faixa_normal": "10.0 a 17.0", "unidade": "%"},
        "PCT": {"nome": "Nível de Procalcitonina", "faixa_normal": "0.10 a 0.28", "unidade": "%"},
        "PLCR": {"nome": "Taxa de Células Plaquetárias Grandes", "faixa_normal": "13.0 a 43.0", "unidade": "%"}
    }

    while True:
        print("\nMenu de Opções:")
        print("1 - Exibir informações de um exame específico")
        print("2 - Exibir exames com valores acima da faixa normal")
        print("3 - Exibir exames com valores abaixo da faixa normal")
        print("4 - Exibir exames dentro da faixa normal")
        print("5 - Listar todos os exames")
        print("6 - Exibir tabela com todos os exames")
        print("7 - Sair")

        opcao = input("Escolha uma opção: ")

        if opcao == "1":
            dado_exame = input("Digite o nome do exame que deseja visualizar as informações: ")
            exibir_informcoes_exames(dado_exame, informacoes_exames)
        elif opcao == "2":
            resultados = buscar_exames_por_faixa(df, informacoes_exames, "acima")
            print("Exames com valores acima da faixa normal:\n", resultados)
        elif opcao == "3":
            resultados = buscar_exames_por_faixa(df, informacoes_exames, "abaixo")
            print("Exames com valores abaixo da faixa normal:\n", resultados)
        elif opcao == "4":
            resultados = buscar_exames_por_faixa(df, informacoes_exames, "nos conformes")
            print("Exames dentro da faixa normal:\n", resultados)
        elif opcao == "5":
            listar_todos_exames(informacoes_exames)
        elif opcao == "6":
            quant_tabela = int(input("\nDigite a quantidade de linhas que deseja visualizar na tabela: "))
            listar_tabela(df, quant_tabela)
        elif opcao == "7":
            print("Saindo do programa.")
            break
        else:
            print("Opção inválida. Tente novamente.")

if __name__ == '__main__':
    main()
