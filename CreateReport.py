import os
import logging
from bs4 import BeautifulSoup
from datetime import datetime
import pandas as pd
import matplotlib.pyplot as plt

class CreateReport():
    def __init__(self, _xlsxPath, _destinyFile, _templatePath):

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
        self.InsertHeaderTitleReport(_title=f"Relatório Gestão Contas")
        self.InsertFooterReport()
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
    def InsertHeaderTitleReport(self, _title):
        self.InsertTextHtmlTag(_tag="td", _text=_title, _id="HeaderTitle")
    
    def InsertFooterReport(self):
        try:
            text = f'Este é um e-mail automático, favor não responder. | Relatório gerado em: ({datetime.now().strftime("%d/%m/%Y")})'
            self.InsertTextHtmlTag(_tag="td", _text=text, _id="Footer")

            self.logger.info(f"[INFO] Text Footer Criado e inserido no relatorio. | CountError: ({self.countErros})")
        except Exception as err:
            self.logger.error(f"[ERROR] Falha ao criar/inserir Texto  Footer no relatorio. | {str(err)}\n")
   
    def InsertPendingExpensesSection(self):
        try:
            def InsertPendingExpensesText():
                dataFrame = self.DespesasPendentes_df[self.DespesasPendentes_df['Situação'] == 'Pendente']
                dataFrame = dataFrame.drop(columns=['Faturamento', 'Data Compra', 'Categoria', 'Descrição', 'Operação', 'Situação'])
                dataFrame = dataFrame.groupby('Responsavel', as_index=False)['Valor'].sum()
                total_pedro = (dataFrame[dataFrame['Responsavel'] == 'Pedro']['Valor']).tolist() # round( , 2 )
                total_outros = (dataFrame[dataFrame['Responsavel'] != 'Pedro']['Valor']).tolist() # round( , 2 )
                text = f"Valor total a ser pago por Pedro = R${round(total_pedro[0] , 2)}, e a soma das despesas para outros responsavéis = R${round(total_outros[0] , 2)}."
                
                self.InsertTextHtmlTag(
                                        _tag = "p", 
                                        _text = text, 
                                        _id = "DespesasPendente"
                                    )
                return
            
            InsertPendingExpensesText()

            def InsertPendingExpensesTable():
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
            
            InsertPendingExpensesTable()

        except Exception as err:
            self.logger.error(f'[ERRO] Falha ao inserir sessao (DespesasPendentes) no relatorio. | {str(err)}\n')
            self.countErros += 1
            return False
