import argparse
from CreateReport import CreateReport
import os

def main(XlsxPath):

    os.system('cls')
    CreateReport(
                    _xlsxPath = XlsxPath,
                    _templatePath = ".\\Templates\\Gestão Contas.html",
                    _destinyFile = "Execution01.html"
                )

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--XlsxPath", type=str, help="Endereço GestaoContas.xlsx")
    args = parser.parse_args()
    main(args.XlsxPath)
