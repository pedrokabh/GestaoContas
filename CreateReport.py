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
        if os.path.isfile(".\\execution.log"):
            os.remove(".\\execution.log")

        logging.basicConfig(filename=f'.\\execution.log', level=logging.INFO, format='%(message)s')
        
        # Variaveis publicas da classe.
        self.logger = logging.getLogger()
        self.countErros = 0 
        self.templatePath = _templatePath
        self.XlsxPath = _xlsxPath
        self.nameReport = _destinyFile
        self.reportDestinyPath = ".\\Execution\\"
        # self.DespesasPagas_df = self.read_excel_sheet( _xlsx_path=_xlsxPath , _sheet_name="Despesas Pagas")
        self.DespesasPendentes_df = self.read_excel_sheet( _xlsx_path=_xlsxPath , _sheet_name="Despesas Pendentes")
        # self.Carteira_df = self.read_excel_sheet( _xlsx_path=_xlsxPath , _sheet_name="Carteira")

        # Criação Relatório.
        self.logger.info(f'[INFO] Iniciando Montagem do Relatorio.')
        self.soup = self.CreateSoupObject()
        self.InsertPendingSection()
        self.SaveSoupAsHtml()
        self.logger.info(f'[INFO][SUCESS] Finalizada Montagem do Relatorio.')

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
            
            return self.logger.info(f'[PATH] Relatorio Pronto em ({file_path}).')

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

    def CreateBarsGraphic(self, _df_EixoX, _df_EixoY, _title=False):
        try:
            self.logger.info("[INFO] Criando grafico de barras...")

            x_pos = range(len(_df_EixoX))

            plt.figure(facecolor='aliceblue')
            plt.bar(x_pos, _df_EixoY, align="center", alpha=0.5)
            plt.xticks(x_pos, _df_EixoX)

            plt.ylabel('R$')
            
            if _title:
                plt.title(_title)

            plt.yticks([])

            ax = plt.gca()

            ax.set_facecolor('aliceblue')

            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['left'].set_visible(False)
            ax.spines['bottom'].set_visible(False)

            for i, v in enumerate(_df_EixoY):
                plt.text(i, v + 0.05 * max(_df_EixoY), str(v), ha='center')
            
            buffer = io.BytesIO()
            plt.savefig(buffer, format='png')
            buffer.seek(0)
            base64_image = base64.b64encode(buffer.read()).decode('utf-8')
            buffer.close()
            
            return base64_image
        
        except Exception as err:
            self.logger.error(f"[ERROR] Falha ao criar grafico de barras. | \n{err}")

    def InsertHtmlTable(self, _soup, _id, _df):
        try:

            table = _soup.find('table', id=_id)
            if table is None:
                raise ValueError(f"[ERROR] Table with id '{_id}' not found")

            tbody = table.find('tbody')
            if tbody is None:
                raise ValueError("[ERROR] Table body not found")

            # Convert DataFrame to list of dictionaries
            data_list = _df.to_dict(orient='records')

            # generating tr
            for row_data in data_list:
                row = _soup.new_tag('tr')

                # generating td's in the tr
                for key in row_data:
                    cell = _soup.new_tag('td')
                    cell.string = str(row_data[key]).replace("00:00:00","")
                    row.append(cell)
                
                tbody.append(row)

            return _soup

        except ValueError as err:
            self.logger.error(f"\n{err}")
            self.countErros += 1
            return False
        except Exception as err:
            self.logger.error(f"[ERROR] Exception Error ao inserir <table id='{_id}'> | \n{err}")
            self.countErros += 1
            return False

    def InsertPendingSection(self):
        try:
            self.logger.info("[INFO] Gerando sessao (DESPESAS PENDENTES)...")
            
            #  1 -- pegando dados das despesas pendentes sem filtros. -- 
            despesasPendentes = self.DespesasPendentes_df

            # 2 -- calculando textos total/responsavel. -- 
            total = despesasPendentes[despesasPendentes['Situação'] == "Pendente"]
            totalPedro = total[total['Responsavel'] == "Pedro"]
            totalOthers = total[total['Responsavel'] != "Pedro"]
            self.InsertTextHtmlTag( _tag = 'h4', _text = f"R$ {str(round(total['Valor'].sum(),2))}",
                                    _id = 'totalPending')
            self.InsertTextHtmlTag( _tag = 'h4', _text = f"R$ {str(round(totalPedro['Valor'].sum(),2))}",
                                    _id = 'totalPedroPending')
            self.InsertTextHtmlTag( _tag = 'h4', _text = f"R$ {str(round(totalOthers['Valor'].sum(),2))}",
                                    _id = 'totalOthersPending')
            del total, totalPedro, totalOthers

            # 3 -- Pegando dados e inserindo na tabela despesa futura -- 
            despesasFuturas = despesasPendentes[despesasPendentes['Situação'] == "Futura"] 
            despesasFuturas = despesasFuturas.drop(columns=['Situação','Responsavel','Data Compra','Operação'])
            despesasFuturas['Faturamento'] = pd.to_datetime(despesasFuturas['Faturamento'], format='%m/%d/%Y')
            despesasFuturas = despesasFuturas[['Descrição', 'Faturamento', 'Categoria', 'Valor']]
            self.soup = self.InsertHtmlTable(
                _soup = self.soup,
                _id = "TablePendingFuture",
                _df = despesasFuturas
            )
            del despesasFuturas
            
            # 4 -- Criando gráfico de barras Operação/Total --
            despesasAtivas = despesasPendentes[despesasPendentes['Situação'] == "Pendente"]
            despesasAtivas = despesasAtivas.groupby('Operação').agg({'Valor': 'sum'}) # Operação se torna index e não uma coluna.
            base64 = 'data:image/png;base64,' + self.CreateBarsGraphic(
                _df_EixoX = despesasAtivas.index.tolist(), # Operação
                _df_EixoY = despesasAtivas['Valor'].tolist() # Total
            )
            self.EditHtmlAtribute(
                _tag="img", 
                _id="BarsGraphicPending", 
                _atributo="src",
                _valorAtributo=f"{base64}"
            )
            del despesasAtivas, base64

            self.logger.info("[INFO] Secao (DESPESAS PENDENTES) inserida com sucesso.")

        except Exception as err:
            self.logger.error(f"[ERROR] Falha ao gerar secao (DESPESAS PENDENTES) | {err}")
            self.countErros += 1
