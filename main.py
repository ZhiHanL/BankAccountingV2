import argparse
import glob

from ExcelMerger import ExcelMerger


DOCUMENTS_PATH = "Transactions/*"


def main():
    arg_parser = argparse.ArgumentParser()
    arg_parser.add_argument('-w', '--workbook')
    args = arg_parser.parse_args()

    statements = glob.glob(DOCUMENTS_PATH)

    ew = ExcelMerger(statements, args.workbook)
    ew.process_transactions()
    print('done :)')


if __name__ == '__main__':
    main()
