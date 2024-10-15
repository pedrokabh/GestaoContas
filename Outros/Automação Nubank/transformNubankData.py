import os
import pandas as pd
import csv

## --- VARIAVEL DE EXECUÇÃO --- #
execution = "Fatura PF"
directory = f"{os.getcwd()}\\Outros\\Automação Nubank\\"
print(f"[Directory] {directory}")

def listar_arquivos_csv(diretorio):
    try:
        arquivos_csv = []

        for arquivo in os.listdir(diretorio):
            if arquivo.endswith('.csv'):
                arquivos_csv.append(arquivo)
        
        return arquivos_csv
    except Exception as err:
        print("Error ListarAequivos -> \n",err)

def definirCategoriaPorDescricao(description, category, operation, value):
    try:
        operation = str(operation)
        description = str(description)
        category = str(category)
        value = str(value)

        if category == "loan_revolving_transitory" and description == "CrÃ©dito de rotativo":
            category = "Juros de Crédito Rotativo"
        elif category == "loan_revolving_tax" or category == "discount_installments" or "iof" in description.lower() or category == "tax_foreign" or "loan_revolving" in category:
            category = "Outras Taxas"

        # Realizar as verificações de contas do mesmo cpf.
        if any(term in description for term in ["42963-0","15445551-3","1001524-3","56681-0","40271-0","410818720-6","52838456-4","32038553-1","18325874-3"]):
              category = "Transferência Entre Contas"
        elif any(term in description for term in ["BANCO BRADESCO S.A"]) and operation == "Boleto":
            operation = "Fatura Cartão Crédito"
            category = "Cartão Crédito (Bradesco)"

        # FILTRO CATEGORIA PELA DESCRIÇÃO
        if any(term in description.lower() for term in ["distribuidora", "bebidas", "bar", "beer", "choperia"]) and "barbosa" not in description.lower():
            category = "Bar/Bebidas"
        elif any(term in description.lower() for term in ["sup epa","pao de acucar","supermercado","villefort"]):
            category = "Supermercado"
        elif any(term in description.lower() for term in ["estacao hookah","thiago luiz januario mendonca", "lucas gabriel alves da silva", "fabricio lene de jesus", "brunna mendes araujo", "carlos junio ferreira da silva", "tabacaria", "ana claudia pereira da silva"]):
            category = "Cigarro"
        elif any(term in description.lower() for term in ["panificadora","vitor mateus assunção camargos", "dona lica", "padaria", "pao divino", "pn trigo vinho", "big mercearia e comerc"]):
            category = "Padaria"
        elif any(term in description.lower() for term in ["zul 10 cartoes","pedro cabrera galindo filho", "estacionamento", "parking", "park", "empresa munic de desenvolvimento urbano rural de bauru"]):
            category = "Estacionamento"
        elif any(term in description.lower() for term in ["walace da silva júnior", "guilherme mendes lima dos santos", "angelo giacomini"]):
            category = "Compras/Catiras"
        elif any(term in description.lower() for term in ["zeut","democrata", "renner", "riachuelo", "c&a"]):
            category = "Vestuario"
        elif "dawidson almeida pires" in description.lower():
            category = "Futebol"
        elif any(term in description.lower() for term in ["pg *ton tec word"]):
            category = "Papelaria"
        elif any(term in description.lower() for term in ["uber", "buser", "motta"]):
            category = "Transporte"
        elif any(term in description.lower() for term in ["conta vivo","telefonica brasil s a", "vivo mg"]):
            category = "VIVO"
        elif any(term in description.lower() for term in ["pag*silmarareisdasilv","barbearia", "silmara reis da silva", "marlon hudson dias cassiano", "igreja evangelica ver"]):
            category = "Barbearia"
        elif any(term in description.lower() for term in ["estado de minas gerais", "receita federal", "secret.", "juros"]):
            category = "Taxas/Imposto"
        elif any(term in description.lower() for term in ["gerais protecao automotiva"]):
            category = "Seguro Carro"
        elif any(term in description.lower() for term in ["3053 conta: 12248-1","18846370-3","escapamentos","mecanica","auto escola","pincel raro papelaria e comercio", "Jrlprdutos", "bh car viaduto","squadra rodas","km auto center bauru","wiliam apolinario moreira","autoonze", "auto pecas", "lava carro", "autopecas"]):
            category = "Carro"
        elif any(term in description.lower() for term in ["bailedaa","sympla", "paymee", "meep"]):
            category = "Lazer"
        elif any(term in description.lower() for term in ["combustivel", "posto", "combustiveis"]):
            category = "Combustível"
        elif any(term in description.lower() for term in ["drogaria", "holiana veloso", "exames"]):
            category = "Saúde"
        elif any(term in description.lower() for term in ["hotmart","xbox","dm *canva","steam","prime video", "amazon prime", "netflix", "cartola"]):
            category = "Assinaturas"
        elif value=="2500,00" and any(term in description.lower() for term in ["24.845.773/0001","realtec solucoes ltda", "isf i negocios ltda", "isf i negocios ltda"]):
            category = "Salário"
        elif value!="2500,00" and any(term in description.lower() for term in ["24.845.773/0001","realtec solucoes ltda", "isf i negocios ltda", "isf i negocios ltda"]):
            category = "Reembolso"
        elif any(term in description.lower() for term in ["fulldarin","frutosdegoias","starbucks","pizza do barulho","churros","garagem hamburgueria","pizza","lanche","breik", "cantina","ifood", "geraldo ferreira de m", "hot dog", "claudioburgue", "bk brasil", "hugo b. dias", "tonolucro", "to no lucro", "churrascaria"]):
            category = "Ifood/Restaurante"
        elif any(term in description for term in ["INSTITUTO CULTURAL NEWTON PAIVA FERREIRA", "CENTRO EDUACIONAL HYARTE ML LTDA"]):
            category = "Educação"
        elif any(term in description for term in ["Leo Gas"]):
            category = "Gas"
        
        
        if category == "transporte":
            category = "Transporte"
        elif category == "restaurante":
            category = "Restaurante"
        elif category == "supermercado":
            category = "Supermercado"
        elif category == "serviÃ§os":
            category = "Serviços"
        elif category == "charge":
            category = "Cobrar"
        elif category == "casa":
            category = "Casa"
        elif category == "lazer":
            category = "Lazer"
        elif category == "outros":
            category = "Outros"
        elif category == "saÃºde":
            category = "Saúde"
        elif category == "eletrÃ´nicos":
            category = "Eletrônicos"
        elif category == "vestuÃ¡rio":
            category = "Vestuario"
        elif category == "educaÃ§Ã£o":
            category = "Educação"

        return category, description, operation, str(value)
    except Exception as err:
        print("Error Filtrar Categoria por Descrição \n",err)

def returnFaturasUnificadasNubank(folder_path):
    try:

        nome_arquivos_Faturas = listar_arquivos_csv(diretorio=folder_path)
        df_faturas_unificadas = []

        for arquivo in nome_arquivos_Faturas:
            print(f"Tratando FATURA - [{str(arquivo)}]")
            dados_arquivos = [] # lista linhas do csv.

            # Extrai todas as linhas do arquivo ATUAL e salva em uma lista.
            with open(str(folder_path+"\\"+arquivo), newline='', encoding='utf-8') as arquivo_csv:
                leitor_csv = csv.reader(arquivo_csv)
                for linha in leitor_csv:
                    dados_arquivos.append(linha)
            
            dados_arquivos.pop(0) # Remove cabeçalho

            for dado in dados_arquivos:
                date = str(dado[0])
                category = str(dado[1])
                description = dado[2]
                value = str(dado[3].replace(".",","))
                operation = None
                movement = None

                if category == "payment" and description == "Pagamento recebido":
                    continue
                else:
                    category, description, operation, value = definirCategoriaPorDescricao(category=category, description=description, operation=operation, value=value)
                    
                    dictonary = {
                        "DATA": date,
                        "VALOR": value,
                        "DESCRIÇÃO": description,
                        "CATEGORIA AUTOMAÇÃO": category,
                        "CATEGORIA SUGERIDA": None,
                        "ARCHIVE": str(arquivo).replace(".csv","")
                    }

                    df_faturas_unificadas.append(dictonary)

        df_faturas = pd.DataFrame(df_faturas_unificadas)
        return df_faturas
    
    except Exception as err:
        print(f"ERROR FATURAS ->\n{err}")

def returnExtratoUnificadoNubank(diretorio_extratos):
    try:
        nome_arquivos_extratos = listar_arquivos_csv(diretorio=diretorio_extratos)
        df_extratos_unificados = []

        for arquivo in nome_arquivos_extratos:
            print(f"Tratando EXTRATOS Arquivo - [{str(arquivo)}]")
            dados_arquivos = [] # lista linhas do csv.

            # Extrai todas as linhas do arquivo ATUAL e salva em uma lista.
            with open(str(diretorio_extratos+"\\"+arquivo), newline='', encoding='utf-8') as arquivo_csv:
                leitor_csv = csv.reader(arquivo_csv)
                for linha in leitor_csv:
                    dados_arquivos.append(linha)
            
            dados_arquivos.pop(0) # Remove cabeçalho

            for dado in dados_arquivos:
                date = str(dado[0]) # DATA
                value = str(dado[1]).replace(".",",") # VALOR
                description = str(dado[3]) # DESCRIÇÃO
                
                # Tratando dados
                operation = description.split('-')[0] # Extraindo operação da descrição.
                description = '-'.join(description.split(' - ')[1:]) # Tirando a operação da STRING, pois já foi extraída anteriormente.
                category = None
                movement = None

                # Filtrando e identificando as RECEITAS.
                if "Crédito em conta" in operation:
                    operation = "Saque Porquinho Nubank"
                    movement = "Receita"
                    category = "Crédito em Conta"
                elif "Estorno" in operation and description != "15445551-3":
                    operation = "Estorno"
                    movement = "Receita"
                    category = "Reembolso"
                elif "Reembolso recebido pelo Pix" in operation:
                    operation = "Estorno"
                    movement = "Receita"
                    category = "Reembolso"
                elif "Transferência Recebida" in operation:
                    operation = "Transferência Recebida"
                    movement = "Receita"
                elif "Transferência recebida pelo Pix" in operation:
                    operation = "Transferência Pix Recebida"
                    movement = "Receita"
                elif "Rendimento Líquido" in description:
                    operation = "Investimento"
                    category = "Aplicação"
                    movement = "Receita"
                    description = "Rendimento Conta Nubank"
                elif "Aplicação RDB" in operation or "Aplicação RDB" in description:
                    operation = "Investimento"
                    category = "Aplicação"
                    description = "Aplicação RDB"
                    movement = "Despesa"
                elif "Resgate RDB" in operation or "Resgate RDB" in description:
                    operation = "Investimento"
                    category = "Aplicação"
                    description = "Resgate RDB"
                    movement = "Receita"
                elif "Rendimento Dinheiro Conta Nubank" in description:
                    operation = "Investimento"
                    category = "Aplicação"
                    description = description
                    movement = "Receita"
                
                # Filtrando e identificando as DESPESAS.
                if "Transferência enviada pelo Pix" in operation:
                    operation = "Pix"
                    movement = "Despesa"
                elif "Pagamento de boleto efetuado" in operation:
                    operation = "Boleto"
                    movement = "Despesa"
                elif "Recarga de celular" in operation:
                    operation = "Recarga Celular"
                    movement = "Despesa"
                    category = "VIVO"
                elif "Pagamento da fatura" in operation or "Pagamento de fatura" in operation:
                    operation = "Fatura Cartão Crédito"
                    category = "Cartão Crédito (Nubank)"
                    movement = "Despesa"
                    description = f"Pagamento da fatura"
                elif "Compra no débito" in operation:
                    operation = "Cartão Débito"
                    movement = "Despesa"
                elif "Compra no débito via NuPay" in operation:
                    operation = "NuPay"
                    movement = "Despesa"


                # Identificando categoria baseado na DESCRIÇÃO.
                category, description, operation, value = definirCategoriaPorDescricao(category=category, description=description, operation=operation, value=value)
                
                dictonary = {
                    "DATA": date,
                    "VALOR": value,
                    "DESCRIÇÃO": description,
                    "CATEGORIA AUTOMAÇÃO": category,
                    "CATEGORIA SUGERIDA": None,
                    "OPERAÇÃO": operation,
                    "MOVIMENTAÇÃO": movement,
                    "ARCHIVE": str(arquivo).replace(".csv","")
                }
                df_extratos_unificados.append(dictonary)
                date = None
                value = None
                description = None

        df_extratos_unificados = pd.DataFrame(df_extratos_unificados)
        return df_extratos_unificados
    
    except Exception as err:
        print(f"ERROR EXTRATOS ->\n{err}")


## -- EXEC -- ##
try:
    if execution == "PF":
        df_extratos = returnExtratoUnificadoNubank(diretorio_extratos=directory)
        df_extratos.to_excel(f"{directory}\ExtratoProcessadoNubankPF.xlsx", index=False)
        # df_extratos.to_csv(f"{directory}\ExtratoProcessadoNubankPF.csv", index=False)
    elif execution == "PJ":
        df_extratos = returnExtratoUnificadoNubank(diretorio_extratos=directory)
        df_extratos.to_excel(f"{directory}\ExtratoProcessadoNubankPJ.xlsx", index=False)
        # df_extratos.to_csv(f"{directory}\ExtratoProcessadoNubankPJ.csv", index=False)
    elif execution == "Fatura PF":
        df_extratos = returnFaturasUnificadasNubank(folder_path=directory)
        df_extratos.to_excel(f"{directory}\FaturaProcessadaNubankPF.xlsx", index=False)
        # df_extratos.to_csv(f"{directory}\FaturaProcessadaNubankPF.csv", index=False)
    else:
        print("Opção não identificada.")

except Exception as err:
    print(f"Error Execution ->\n{err}")
    
## -- -- FIM EXECUÇÃO -- -- ##
