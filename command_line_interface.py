#!/usr/bin/python
import no_more_excel
import argparse


parser = argparse.ArgumentParser(prog='No More Excel')

parser.add_argument('mode', choices=['json_to_xlsx', 'xlsx_to_json'], help='json_to_xlsx or xlsx_to_json')
parser.add_argument('start-index', help='index of the first cell of the parsed xlsx zone')
parser.add_argument('parsed-zone', type=int, nargs=2, help='two integer that define respectively the number of column and the number of line of the parsed zone')
parser.add_argument('input', help='file taht will be parsed')
parser.add_argument('output', help='output file')
parser.add_argument('duplicate-from', default=None, help='if you want your xlsx output file duplucated from an existing one, put his name in this argument')

args = parser.parse_args()

if args.mode == 'json_to_xlsx':
    no_more_excel.json_to_xlsx( duplicate_from=args.duplicate_from,
                                ini=args.input,
                                out=args.output,
                                start_index=args.start_index,
                                nb_columns=args.parsed_zone[0],
                                nb_items=args.parsed_zone[1])
elif args.mode == 'xlsx_to_json':
    no_more_excel.xlsx_to_json( ini=args.input,
                                out=args.output,
                                start_index=args.start_index,
                                nb_columns=args.parsed_zone[0],
                                nb_items=args.parsed_zone[1])
