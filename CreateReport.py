import os
import logging
from bs4 import BeautifulSoup
from datetime import datetime

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
        
        # Criação Relatório.
        self.logger.info('')
        self.logger.info(f'[INFO] Iniciando Montagem do Relatorio.')
        self.soup = self.CreateSoupObject()
        self.InsertTitleReport(_title=f"Relatório Gestão Contas")
        self.InsertFooterReport()
        self.SaveSoupAsHtml()

        return
    
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
        
    # -- Editando HTML -- #
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
    def InsertTitleReport(self, _title):
        self.InsertTextHtmlTag(_tag="td", _text=_title, _id="TitleReport")
    
    def InsertFooterReport(self):
        try:
            text = f'Este é um e-mail automático, favor não responder. | Relatório gerado em: ({datetime.now().strftime("%d/%m/%Y")})'
            self.InsertTextHtmlTag(_tag="td", _text=text, _id="Footer")

            self.logger.info(f"[INFO] Text Footer Criado e inserido no relatorio. | CountError: ({self.countErros})")
        except Exception as err:
            self.logger.error(f"[ERROR] Falha ao criar/inserir Texto  Footer no relatorio. | {str(err)}\n")

    def NubankCreditoSection(self):
        try:
            text = "Este é um e-mail automático, favor não responder. | Relatório gerado em: (05/03/2024)"
            self.InsertTextHtmlTag(_tag="td", _text=text, _id="Footer")

            self.logger.info(f"[INFO] Sessão Nubank Criada e inserida no relatorio. | CountError: ({self.countErros})")
        except Exception as err:
            self.logger.error(f"[ERROR] Falha ao criar/inserir sessao nubank no relatorio. | ")
        