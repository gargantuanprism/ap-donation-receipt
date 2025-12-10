#!/usr/bin/env python

import argparse
from string import Template
import csv
import datetime
import sys
import os

from babel.dates import format_date
import dateutil

BASE_VALUES = {
  'sponsor_name': os.getenv('SPONSOR_NAME'),
  'org_name': os.getenv('ORG_NAME'),
  'org_addr': os.getenv('ORG_ADDR'),
  'org_ein': os.getenv('ORG_EIN'),
  'authorized_agent_name': os.getenv('AGENT_NAME'),
  'authorized_agent_title': os.getenv('AGENT_TITLE'),
  'letter_date': format_date(datetime.date.today(), locale='en_US'),
  'sig_image': os.getenv('SIG_IMAGE'),
  'org_logo': os.getenv('ORG_LOGO'),
  'sponsor_logo': os.getenv('SPONSOR_LOGO')
}

ZEFFY_START_LINE = 2
ZEFFY_COLS = {
  'date': 0,
  'amount': 1,
  'first_name': 5,
  'last_name': 6,
  'email': 7
}

BENEVITY_START_LINE = 12
BENEVITY_COLS = {
  'date': 2,
  'amount': 19,
  'first_name': 3,
  'last_name': 4,
  'email': 5
}

def get_user_data(row, args):
  cols = ZEFFY_COLS if args.zeffy else BENEVITY_COLS

  if args.benevity:
    donation_amount = row[cols['amount']] if len(row) >= 19 else row[12]
  else:
    donation_amount = row[cols['amount']]

  if not donation_amount.startswith('$'):
    donation_amount = f'${donation_amount}'

  donation_date = dateutil.parser.parse(row[cols['date']])

  data = {
    'donation_date': donation_date.strftime('%Y/%m/%d'),
    'donation_amount': donation_amount,
    'first_name': row[cols['first_name']].title(),
    'last_name': row[cols['last_name']].title(),
    'donor_email': row[cols['email']],
    'donation_desc': '',
    'nc_value': ''
  }

  if args.donation_desc:
    data['donation_desc'] = args.donation_desc

  return data

def write_files(tpl_data, args):
  date = tpl_data['donation_date'].replace('/', '-')
  report_type = 'Benevity' if args.benevity else 'Zeffy'
  outfile_basename = f"APB_donation_receipt_{report_type}_{date}_{tpl_data['first_name']}_{tpl_data['last_name']}"

  with open(f'out/{outfile_basename}.md', 'w') as out_f:
    out_f.write(tpl.substitute(**tpl_data))

  if args.pdf:
    os.system(f'pandoc out/{outfile_basename}.md -o out/{outfile_basename}.pdf')

  if args.html:
    os.system(f'pandoc out/{outfile_basename}.md -o out/{outfile_basename}.html')

  return f'out/{outfile_basename}'

if __name__ == '__main__':
  parser = argparse.ArgumentParser()
  parser.add_argument('filename', nargs='+')
  parser.add_argument('--benevity', action='store_true')
  parser.add_argument('--zeffy', action='store_true')
  parser.add_argument('--donation-desc')
  parser.add_argument('--pdf', action='store_true')
  parser.add_argument('--html', action='store_true')
  parser.add_argument('--first-name')
  parser.add_argument('--last-name')
  parser.add_argument('--email')
  parser.add_argument('--amount', type=float)
  parser.add_argument('--date')
  parser.add_argument('--nc-value', type=float)
  args = parser.parse_args()

  with open('template.md', 'r') as f:
    tpl = Template(f.read())

  for filename in args.filename:
    with open(filename, 'r') as f:
      if args.benevity:
        for _ in range(BENEVITY_START_LINE):
          next(f)

        reader = csv.reader(f)
        for row in reader:
          if row[0] == 'Totals':
            break

          user_data = get_user_data(row, args)
          tpl_data = BASE_VALUES | user_data
          outfile_basename = write_files(tpl_data, args)
          print(f'{tpl_data["donor_email"]} -> {outfile_basename}')

      elif args.zeffy:
        for _ in range(ZEFFY_START_LINE):
          next(f)

        reader = csv.reader(f)
        for row in reader:
          user_data = get_user_data(row, args)
          tpl_data = BASE_VALUES | user_data
          outfile_basename = write_files(tpl_data, args)
          print(f'{tpl_data["donor_email"]} -> {outfile_basename}')

      # manual
      else:
        args.zeffy = True

        row = [
          args.date,
          args.amount,
          None,
          None,
          None,
          args.first_name,
          args.last_name,
          args.email
        ]

        tpl_data = BASE_VALUES | get_user_data(row, args)
        print(tpl_data)

