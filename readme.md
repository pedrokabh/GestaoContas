# GestaoContas
Conjunto de arquivos para análise de extratos e faturas de cartões de crédito. Nestes arquivos, existem scripts para tratamento de dados e criação de arquivos csv/xlsx estruturados para construção de relatórios em PowerBI, apartir destes arquivos gerados. Utilizo estes arquivos para gerenciamento e planejamento financeiro pessoal.

* Dashboard Despesas Pendetes / Extraído do Arquivo 'Despesas Pendentes.xlsx'.

![Captura de tela](./Imagens%20ReadMe/DashboardDespesasPendentes.png)

* Dashboard Análise de Extratos / Extraído de Arquivos de Extratos emitidos pelas instituições financeiras.

![Captura de tela](./Imagens%20ReadMe/DashboardExtratos.png)

* Dashboard Análise de Extratos / Extraído dos arquivos extratos e faturas emitidos pelas instituições financeiras.

![Captura de tela](./Imagens%20ReadMe/DashboardHistoricoDespesas.png)

## Script para tratar arquivos de faturas e extratos gerados pelo Nubank.

### Configuração do Script:
o script 'GestaoContas\Dados de Contas\Dados Nubank\Automação Nubank\DespesasProcessadasNubank.py' pode resultar em três arquivos caso executado corretamente:
* FaturaProcessadaNubankPF.csv.
* ExtratoProcessadoNubankPJ.csv.
* ExtratoProcessadoNubankPF.csv.

Para isso deve-se escolher o valor do parâmetro de execução execution ["PF","PJ","Fatura PF"]: 
```Python
    ## -- EXEC -- ##
    try:
        if execution == "PF":
            directory = r".\Executar PF"
            df_extratos = returnExtratoUnificadoNubank(diretorio_extratos=directory)
            df_extratos.to_csv(r".\ExtratoProcessadoNubankPF.csv", index=False)
        elif execution == "PJ":
            directory = r".\Executar PJ"
            df_extratos = returnExtratoUnificadoNubank(diretorio_extratos=directory)
            df_extratos.to_csv(r".\ExtratoProcessadoNubankPJ.csv", index=False)
        elif execution == "Fatura PF":
            directory = r".\Fatura PF"
            df_extratos = returnFaturasUnificadasNubank(folder_path=directory)
            df_extratos.to_csv(r".\FaturaProcessadaNubankPF.csv", index=False)
        else:
            print("Opção não identificada.")

    except Exception as err:
        print(f"Error Execution ->\n{err}")
        
    ## -- -- FIM EXECUÇÃO -- -- ##
```