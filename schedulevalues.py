import pandas as pd
import numpy as np
import openpyxl
import os

def main():
  filefound = False
  svdf = pd.DataFrame()
  df = pd.DataFrame()
  generate = True
  while not filefound:
    try:
      street = input("What is the name of the street? ")
      street = street.lower()
      street = street.replace(" ", "")
      num = input("What is the street number? ")
      lotnum = input("What is the lot number? ")
      print(f'searching for {num}{street}lot{lotnum}.csv')
      df = pd.read_csv(f'{num}{street}lot{lotnum}.csv')
      svdf = pd.read_csv(f'{num}{street}lot{lotnum}sv.csv')
      # Clean 'Percent Billed' column by replacing NA or inf values with 0
      svdf[f'Lot {lotnum}'] = svdf[f'Lot {lotnum}'].astype('string')
      svdf['Scheduled Value'] = svdf['Scheduled Value'].astype(float)
      svdf['Previous Period'] = svdf['Previous Period'].astype(float)
      svdf['This Period'] = svdf['This Period'].astype(float)
      svdf['Total Billed'] = svdf['Total Billed'].astype(float)
      svdf['Percent Billed'] = svdf['Percent Billed'].astype(int)
      svdf['Balance to Finish'] = svdf['Balance to Finish'].astype(float)
      svdf['Percent Billed'].fillna(0, inplace=True)
      filefound = True
      generate = False
      
    except FileNotFoundError:
      choice = input("File not found. Would you like to generate a new schedule of values? Y/N ")
      choice = choice.lower()
      while not choice.startswith('n') and not choice.startswith('y'):
        choice = input('Input not recognized. Would you like to generate a new schedule of values? Y/N')
      # generate new schedule of values
      if choice.startswith('n'):
        filefound = False
        continue
      elif choice.startswith('y'):
        data = {
          f'Lot {lotnum}': ['Concrete', 'Masonry and Siding', 'Soffit', 'Electric Install', 'Electric Fixtures', 'Millwork & Trim Includes', 'Cabinets and Vanities', 'Finish Paint', 'Flooring', 'Overhead/profit', 'Other', 'Total'],
          'Scheduled Value': [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
          'Previous Period': [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
          'This Period': [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
          'Total Billed': [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
          'Percent Billed': [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
          'Balance to Finish': [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        }
        svdf = pd.DataFrame(data)
        print('New schedule of values generated.')
        print(svdf)
        svdf = setBudget(svdf, lotnum)
        filefound = True
        get_totals(svdf, True)
  
  
  svdf[f'Lot {lotnum}'] = svdf[f'Lot {lotnum}'].astype('string')
  svdf['Scheduled Value'] = svdf['Scheduled Value'].astype(float)
  svdf['Previous Period'] = svdf['Previous Period'].astype(float)
  svdf['This Period'] = svdf['This Period'].astype(float)
  svdf['Total Billed'] = svdf['Total Billed'].astype(float)
  svdf['Percent Billed'] = svdf['Percent Billed'].astype(int)
  svdf['Balance to Finish'] = svdf['Balance to Finish'].astype(float)
  input()

  svdf = get_values(df, svdf, lotnum)
  input('get values complete')
  svdf = update_schedule(svdf, lotnum)
  
  svdf = get_totals(svdf, False)
  print(svdf)
  svdf.to_csv(f'{num}{street}lot{lotnum}sv.csv', index=False)
  input('Press Enter to complete')


def get_values(df, svdf, lotnum):
  input()
  for i, j in zip(svdf.loc[:, 'This Period'], svdf.loc[:, 'Previous Period']):
    j += i
    i = 0
  for i, j in zip(df.loc[:, 'Code'], df.loc[:, 'Amount']):
    for k in svdf[f'Lot {lotnum}']:
      if str(i).lower() in str(k).lower():
        svdf.loc[svdf[f'Lot {lotnum}'] == str(k), 'This Period'] += j
        print(svdf.loc[svdf[f'Lot {lotnum}'] == str(k), :])
  return svdf

def update_schedule(svdf, lotnum):
  
  for i in svdf[f'Lot {lotnum}']:
    total = svdf.loc[svdf[f'Lot {lotnum}'] == str(i), 'This Period'] + svdf.loc[svdf[f'Lot {lotnum}'] == str(i), 'Previous Period']
    svdf.loc[svdf[f'Lot {lotnum}'] == str(i), 'Total Billed'] = round(total, 2)
    budget = svdf.loc[svdf[f'Lot {lotnum}'] == str(i), 'Scheduled Value']
    percent = round(100 * (total / budget))
    svdf.loc[svdf[f'Lot {lotnum}'] == str(i), 'Percent Billed'] = percent
    svdf.loc[svdf[f'Lot {lotnum}'] == str(i), 'Balance to Finish'] = round(budget - total, 2)
  return svdf

def get_totals(svdf, initial):
  totalbudget = 0
  totalprev = 0
  totalthis = 0
  totalbill = 0
  for i, j, k, l in zip(svdf['Scheduled Value'], svdf['Previous Period'], svdf['This Period'], svdf['Total Billed']):
    totalbudget += i
    totalprev += j
    totalthis += k
    totalbill += l
  if initial == False:
    svdf.loc[11, 'Scheduled Value'] = round(totalbudget, 2)
  svdf.loc[11, 'Previous Period'] = round(totalprev, 2)
  svdf.loc[11, 'This Period'] = round(totalthis, 2)
  svdf.loc[11, 'Total Billed'] = round(totalbill, 2)
  svdf.loc[11, 'Percent Billed'] = round(100 * (totalbill / totalbudget))
  svdf.loc[11, 'Balance to Finish'] = round(totalbudget - totalbill, 2)
  return svdf

def setBudget(svdf, lotnum):
  for i in svdf[f'Lot {lotnum}']:
    if i == 'Total':
      continue
    budget = input(f'What is the budget for {i}? ')
    budget = float(budget)
    svdf.loc[svdf[f'Lot {lotnum}'] == str(i), 'Scheduled Value'] = budget
  svdf = get_totals(svdf, True)
  return svdf
  
if __name__ == '__main__':
  main()
