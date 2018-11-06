#!/usr/bin/python
import no_more_excel
import argparse


parser = argparse.ArgumentParser(description='No More Excel')

parser.add_argument('mode', metavar='N', type=string,choices=['json_to_xlsx', 'xlsx_to_json'] help='json_to_xlsx or xlsx_to_json')




json_to_xlsx(duplicate,ini="in.json",out="out.xlsx",start_index="A0",nb_columns=0,nb_items=0):
xlsx_to_json(ini="in.xlsx",out="out.json",start_index="A0",nb_columns=0,nb_items=0):
