#!/usr/bin/python
from no_more_excel import json_to_xlsx, xlsx_to_json
import argparse


parser = argparse.ArgumentParser(prog='command_line_interface.py')
subparsers = parser.add_subparsers()

parser_json_to_xlsx = subparsers.add_parser('json-to-xlsx', help='convert json to a xlsx targetted zone (see json_to_xlsx -h)')
parser_json_to_xlsx.set_defaults(func=json_to_xlsx)
parser_json_to_xlsx.add_argument('--input', '-i', required=True, help='json file\'s name taht will be parsed')
parser_json_to_xlsx.add_argument('--output', '-o', required=True, help='xlsx output file name')
parser_json_to_xlsx.add_argument('--duplicate-from', '-d', required=True,default=None, type=str, help='if you want your xlsx output file duplucated from an existing one, put his name in this argument')
parser_json_to_xlsx.add_argument('--start-index', '-s', required=True, metavar='[A-Z]+[0-9]+', help='index of the first cell of the targetted xlsx zone')
parser_json_to_xlsx.add_argument('--parsed-zone', '-z', required=True, type=int, nargs=2, metavar=('X','Y'), help='two integers that defines respectively the number of column and the number of line of the targetted zone')

parser_xlsx_to_json = subparsers.add_parser('xlsx-to-json', help='convert a xlsx targetted zone to json (see xlsx_to_json -h)')
parser_xlsx_to_json.set_defaults(func=xlsx_to_json)
parser_xlsx_to_json.add_argument('--input', '-i', required=True, help='xlsx file\'s name taht will be parsed')
parser_xlsx_to_json.add_argument('--output', '-o', required=True, help='json output file name')
parser_xlsx_to_json.add_argument('--start-index', '-s', metavar='[A-Z]+[0-9]+', required=True, help='index of the first cell of the parsed xlsx zone')
parser_xlsx_to_json.add_argument('--parsed-zone', '-z', type=int, nargs=2, metavar=('X','Y'), required=True, help='two integers that defines respectively the number of column and the number of line of the parsed zone')

args = parser.parse_args()

if args.func is xlsx_to_json:
    xlsx_to_json( ini=args.input,
                  out=args.output,
                  start_index=args.start_index,
                  nb_columns=args.parsed_zone[0],
                  nb_items=args.parsed_zone[1])

elif args.func is json_to_xlsx:
    json_to_xlsx( duplicate_from=args.duplicate_from,
                  ini=args.input,
                  out=args.output,
                  start_index=args.start_index,
                  nb_columns=args.parsed_zone[0],
                  nb_items=args.parsed_zone[1])
