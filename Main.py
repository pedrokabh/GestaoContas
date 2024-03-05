import argparse
from CreateReport import CreateReport
import os

def main(XlsxPath):

    # print(f'Arquivo (Gestão Contas.xlsx) ({XlsxPath})')
    os.system('cls')
    CreateReport(_xlsxPath = XlsxPath)

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--XlsxPath", type=str, help="Endereço GestaoContas.xlsx")
    args = parser.parse_args()
    main(args.XlsxPath)
