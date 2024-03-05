import os
import logging
from bs4 import BeautifulSoup 
from datetime import datetime
import pandas as pd
import matplotlib.pyplot as plt
import io
import base64

# COMANDO PIP PARA INSTALAR AS BIBLIOTECAS.
# pip install matplotlib beautifulsoup4 pandas openpyxl

class CreateReport():
    def __init__(self, _xlsxPath, _destinyFile, _templatePath):

        # ! TESTE ! #
        if os.path.isfile(".\\Another Archives\\execution.log"):
            print('LOG REMOVIDO!\n')
            os.remove(".\\Another Archives\\execution.log")

        logging.basicConfig(filename=f'.\\Another Archives\\execution.log', level=logging.INFO, format='%(message)s')
        
        # Variaveis publicas da classe.
        self.logger = logging.getLogger()
        self.countErros = 0 
        self.templatePath = _templatePath
        self.XlsxPath = _xlsxPath
        self.nameReport = _destinyFile
        self.reportDestinyPath = ".\\Generated Reports\\"
        self.DespesasPagas_df = self.read_excel_sheet( _xlsx_path=_xlsxPath , _sheet_name="Despesas Pagas")
        self.DespesasPendentes_df = self.read_excel_sheet( _xlsx_path=_xlsxPath , _sheet_name="Despesas Pendentes")
        self.Carteira_df = self.read_excel_sheet( _xlsx_path=_xlsxPath , _sheet_name="Carteira")

        # Criação Relatório.
        self.logger.info('')
        self.logger.info(f'[INFO] Iniciando Montagem do Relatorio.')
        self.soup = self.CreateSoupObject()
        self.InsertTextHtmlTag(_tag="td", _text=f"Relatório Gestão Contas", _id="HeaderTitle") # Inserindo título de cabeçalho.
        self.InsertTextHtmlTag(_tag="td", _text=f'Este é um e-mail automático, favor não responder. | Relatório gerado em: ({datetime.now().strftime("%d/%m/%Y")})', _id="Footer") # Adicionando mensagem de rodapé.
        self.InsertPendingExpensesSection()
        self.SaveSoupAsHtml()

        return
    
    # -- Pandas -- #
    def read_excel_sheet(self, _xlsx_path, _sheet_name):
        try:
            df = pd.read_excel(_xlsx_path, sheet_name=_sheet_name)
            self.logger.info(f"[INFO] Data Frame ({_sheet_name}) Criado com sucesso.")
            return df
        except Exception as e:
            self.logger.error(f"[ERRO] Falha ao gerar data frame. Arquivo ({_xlsx_path}) SheetName ({_sheet_name}) | {str(e)}\n")
            return False
        
    # -- Utillizando biblioteca BeatifulSoup --
    def CreateSoupObject(self):
        try:
            with open(self.templatePath, "r", encoding="utf-8") as file:
                    html_content = file.read()
            soup = BeautifulSoup(html_content, "html.parser")

            self.logger.info(f'[INFO] Criado Objeto BeatifulSoup.')
            return soup
        
        except Exception as err:
            self.logger.error(f"[ERRO] Falha ao criar objeto BeatifulSoup. | {str(err)}\n")
            self.countErros += 1
            return False
    
    def SaveSoupAsHtml(self):
        try:
            destiny_directory = os.path.dirname(self.reportDestinyPath)
            file_path = os.path.join(self.reportDestinyPath, self.nameReport) 

            if not os.path.exists(destiny_directory): # Verificando se o diretório existe e criando-o se não existir
                os.makedirs(destiny_directory)

            with open(file_path, "w", encoding="utf-8") as file: # Escrevendo o relatório modificado no arquivo HTML
                file.write(str(self.soup))
            
            return self.logger.info(f'[INFO] Relatorio Pronto em ({file_path}).')

        except Exception as err:
            self.logger.error(f"[ERROR] falha ao salvar BeatifulSoup como arquivo.html | {err}\n")
            self.countErros += 1
            return False
        
    def EditHtmlAtribute(self, _tag, _id, _atributo, _valorAtributo):
        try:
            elemento = self.soup.find(_tag, id=_id)
            if elemento:
                elemento[_atributo] = _valorAtributo

                self.logger.info(f"[INFO] Atributo '{_atributo}' alterado com sucesso. | <{_tag} id='{_id}'>")
                return self.soup
            
        except Exception as err:
            self.logger.error(f"[ERROR] Falha ao alterar atributo '{_atributo}' na Tag <{_tag} id='{_id}'> | {str(err)}\n")
            self.countErros += 1
            return False

    def InsertTextHtmlTag(self, _tag, _text, _id):
        try:
            target_element = self.soup.find(_tag, {"id": _id})
            if target_element:
                target_element.string = _text
                self.logger.info(f"[INFO] Texto inserido na tag com sucesso. | <{_tag} id='{_id}'>")
            else:
                self.logger.warning(f"[WARNING] Tag não encontrada <{_tag} id='{_id}'> | Texto não inserido na tag.")
                self.countErros += 1
                return False
            return self.soup

        except Exception as err:
            self.logger.error(f"[ERROR] Falha ao inserir texto na tag html | {str(err)}\n")
            self.countErros += 1
            return False
   
   # -- Criando Seções do Relatório -- #
    def InsertPendingExpensesSection(self):
        try:
            # -- TEXTO DA SEÇÃO -- #
            def InsertSectionText():
                dataFrame = self.DespesasPendentes_df[self.DespesasPendentes_df['Situação'] == 'Pendente']
                dataFrame = dataFrame.drop(columns=['Faturamento', 'Data Compra', 'Categoria', 'Descrição', 'Operação', 'Situação'])
                dataFrame = dataFrame.groupby('Responsavel', as_index=False)['Valor'].sum()
                total_pedro = (dataFrame[dataFrame['Responsavel'] == 'Pedro']['Valor']).tolist()
                total_outros = (dataFrame[dataFrame['Responsavel'] != 'Pedro']['Valor']).tolist()
             
                self.InsertTextHtmlTag(
                                        _tag = "p", 
                                        _text = f"VALOR TOTAL A SER PAGO: R${dataFrame['Valor'].sum()}.", 
                                        _id = "DespesasPendente"
                                    )
                
                self.InsertTextHtmlTag(
                                        _tag = "p", 
                                        _text = f'VALOR TOTAL A SER PAGO PEDRO: R${round(sum(total_pedro), 2)}.', 
                                        _id = "DespesasPendente1"
                                    )
                
                self.InsertTextHtmlTag(
                                    _tag='p', 
                                    _text=f'VALOR TOTAL A SER PAGO TERCEIROS: R${round(sum(total_outros), 2)}.', 
                                    _id='DespesasPendente2'
                                )
                return
            InsertSectionText()

            # -- TABELA DA SEÇÃO -- #
            def InsertSectionTable():
                # Filtra o DataFrame da classe
                dataFrame = self.DespesasPendentes_df[self.DespesasPendentes_df['Situação'] == 'Pendente']
                dataFrame = dataFrame.drop(columns=['Situação', 'Descrição', 'Faturamento', 'Data Compra'])
                dataFrame = dataFrame.groupby(['Categoria', 'Operação', 'Responsavel'], as_index=False)['Valor'].sum()

                # Encontra a tag tbody onde os dados serão inseridos
                tbody_tag = self.soup.find('tbody', {'id': 'TablePendencias'})
                # Limpa qualquer conteúdo existente na tbody
                tbody_tag.clear()

                # Loop pelos dados do DataFrame e insere-os na tabela HTML
                for index, row in dataFrame.iterrows():
                    tr_tag = self.soup.new_tag('tr')
                    tbody_tag.append(tr_tag)

                    # Adiciona os valores de Categoria, Operação, Valor e Responsável a cada tr
                    for column in ['Categoria', 'Operação', 'Valor', 'Responsavel']:
                        td_tag = self.soup.new_tag('td')
                        td_tag.string = str(row[column])
                        tr_tag.append(td_tag)

                return self.soup
            InsertSectionTable()

            # -- GRÁFICO DE DESPESA POR OPERAÇÕES -- #
            def InsertBarsGraphic():
                try:
                    dataFrame = self.DespesasPendentes_df[self.DespesasPendentes_df['Situação']=='Pendente']
                    dataFrame = dataFrame.drop(columns=['Faturamento', 'Data Compra', 'Categoria', 'Descrição', 'Responsavel', 'Situação'])
                    dataFrame = dataFrame.groupby(['Operação'], as_index=False)['Valor'].sum()
                    
                    Operacao = dataFrame['Operação'].tolist() # Dados para o gráfico de barras
                    Valor = dataFrame['Valor'].tolist()
                    cores = []
                    for op in Operacao:
                        if op == 'Nubank Crédito':
                            cores.append('#2d69d9')
                        elif op == 'Pix Nubank':
                            cores.append('#4c81e3')
                        elif op == 'Débito Nubank':
                            cores.append('#6a98ec') 
                        elif op == 'Boleto Nubank':
                            cores.append('#89b0f6') 
                        else:
                            cores.append('#a7c7ff')  

                    plt.bar(Operacao, Valor, color=cores, linewidth=0) # Criar o gráfico de barras com cores definidas
                    plt.title('MÉTODO PAGAMENTO') # Adicionar título e rótulos aos eixos
                    ax = plt.gca() # Desativando Bordas do Gráfico
                    ax.spines['top'].set_visible(False)
                    ax.spines['right'].set_visible(False)
                    ax.spines['bottom'].set_visible(False)
                    ax.spines['left'].set_visible(False)
                    plt.yticks([]) # Removendo legenda eixo y
                    
                    for i in range(len(Operacao)): # Adicionando os valores acima das barras
                        plt.text(Operacao[i], Valor[i], str(round(Valor[i],2)), ha='center', va='bottom')

                    temp_file = io.BytesIO() # Salvar o gráfico temporariamente
                    plt.savefig(temp_file, format='png')
                    temp_file.seek(0)

                    # plt.show() # Exibir o gráfico
                    base64_encoding = base64.b64encode(temp_file.read()).decode('utf-8')
                    plt.close() # Fechar o gráfico para liberar recursos

                    self.soup = self.EditHtmlAtribute( # Adiciona gráfico no template.
                            _tag="img",
                            _id="OperationsGraphicPendingExpenses",
                            _atributo="src",
                            _valorAtributo=f"data:image/png;base64,{base64_encoding}",
                        )

                    return True
                except Exception as err:
                    self.logger.error(f'[ERRO] Falha ao criar grafico despesas por operação (DespesasPendentes) | {str(err)}\n')
                    self.countErros += 1
                    return False
            InsertBarsGraphic()

            # -- GRÁFICO DE PIZZA DESPESA POR CATEGORIA -- #
            def InsertPizzaBarGraphic():
                try:
                    dataFrame = self.DespesasPendentes_df[self.DespesasPendentes_df['Situação'] == 'Pendente']
                    dataFrame = dataFrame.drop(columns=['Faturamento', 'Data Compra', 'Responsavel', 'Descrição', 'Operação', 'Situação'])
                    dataFrame = dataFrame.groupby(['Categoria'], as_index=False)['Valor'].sum()

                    labels = dataFrame['Categoria'].tolist() # Dados para o gráfico de pizza
                    quantidades = dataFrame['Valor'].tolist()

                    plt.pie(quantidades, labels=labels, autopct='%1.1f%%') # Criar o gráfico de pizza
                    plt.title('CATEGORIAS') # Adicionar um título
                    
                    temp_file = io.BytesIO() # Salvar o gráfico temporariamente
                    plt.savefig(temp_file, format='png')
                    temp_file.seek(0)

                    # plt.show() # Exibir o gráfico
                    base64_encoding = base64.b64encode(temp_file.read()).decode('utf-8')
                    plt.close() # Fechar o gráfico para liberar recursos

                    self.soup = self.EditHtmlAtribute( # Adiciona gráfico no template.
                            _tag="img",
                            _id="PizzaGraphicPendingExpenses",
                            _atributo="src",
                            _valorAtributo=f"data:image/png;base64,{base64_encoding}",
                        )
                    
                    self.logger.info('[INFO] Grafico de Barras Criado/Inserido (PendingExpensesSection)')
                    return True

                except Exception as err:
                    self.logger.error(f'[ERROR] Falha ao criar/inserir grafico de pizza. (PendingExpensesSection) | {err}\n')
                    self.countErros += 1
                    return False
            InsertPizzaBarGraphic()

        except Exception as err:
            self.logger.error(f'[ERRO] Falha ao inserir sessao (DespesasPendentes) no relatorio. | {str(err)}\n')
            self.countErros += 1
            return False
