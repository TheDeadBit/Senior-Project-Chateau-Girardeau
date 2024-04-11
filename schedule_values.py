import pandas as pd


def main():
  file_found = False
  svdf = pd.DataFrame()
  df = pd.DataFrame()

  while not file_found:
    try:
      street_name = input("What is the name of the street? ")
      street_name = street_name.lower()
      street_name = street_name.replace(" ", "")
      street_number = input("What is the street number? ")
      house_number = input("What is the House Number? ")

      print(f'searching for {street_number}{street_name}lot{house_number}.csv')

      df = pd.read_csv(f'{street_number}{street_name}lot{house_number}.csv')
      svdf = pd.read_csv(f'{street_number}{street_name}lot{house_number}sv.csv')
      
      # Clean 'Percent Billed' column by replacing NA or inf values with 0
      svdf[f'House Number {house_number}'] = svdf[f'House Number {house_number}'].astype('string')
      svdf['Scheduled Value'] = svdf['Scheduled Value'].astype(float)
      svdf['Previous Period'] = svdf['Previous Period'].astype(float)
      svdf['This Period'] = svdf['This Period'].astype(float)
      svdf['Total Billed'] = svdf['Total Billed'].astype(float)
      svdf['Percent Billed'] = svdf['Percent Billed'].astype(int)
      svdf['Balance to Finish'] = svdf['Balance to Finish'].astype(float)
      svdf['Percent Billed'].fillna(0, inplace=True)
      
      
    except FileNotFoundError:
      choice = input("File not found. Would you like to generate a new schedule of values? Y/N ")
      choice = choice.lower()
      
      while not choice == 'n' and not choice == 'y':
        choice = input('Input not recognized. Would you like to generate a new schedule of values? Y/N')
      # generate new schedule of values
      
      if choice.startswith('n'):
        continue
      
      elif choice.startswith('y'):
        data = {
          f'House Number {house_number}': ['Concrete', 'Masonry and Siding', 'Soffit', 'Electric Install', 'Electric Fixtures', 'Millwork & Trim Includes', 'Cabinets and Vanities', 'Finish Paint', 'Flooring', 'Overhead/profit', 'Other', 'Total'],
          'Scheduled Value': [0 for _ in range(12)],
          'Previous Period': [0 for _ in range(12)],
          'This Period': [0 for _ in range(12)],
          'Total Billed': [0 for _ in range(12)],
          'Percent Billed': [0 for _ in range(12)],
          'Balance to Finish': [0 for _ in range(12)]
        }

        svdf = pd.DataFrame(data)
        print('New schedule of values generated.')
        print(svdf)
        svdf = set_budget(svdf, house_number)
        get_totals(svdf, True)
    
    file_found = True
  
  
  svdf[f'House Number {house_number}'] = svdf[f'House Number {house_number}'].astype('string')
  svdf['Scheduled Value'] = svdf['Scheduled Value'].astype(float)
  svdf['Previous Period'] = svdf['Previous Period'].astype(float)
  svdf['This Period'] = svdf['This Period'].astype(float)
  svdf['Total Billed'] = svdf['Total Billed'].astype(float)
  svdf['Percent Billed'] = svdf['Percent Billed'].astype(int)
  svdf['Balance to Finish'] = svdf['Balance to Finish'].astype(float)
  print()

  svdf = get_values(df, svdf, house_number)
  print('get values complete')
  svdf = update_schedule(svdf, house_number)
  
  svdf = get_totals(svdf, False)
  print(svdf)
  svdf.to_csv(f'{street_number}{street_name}House Number {house_number}sv.csv', index=False)
  input('Press Enter to complete')


def get_values(df, svdf, house_number):
  print()
  for i, j in zip(svdf.loc[:, 'This Period'], svdf.loc[:, 'Previous Period']):
    j += i
    i = 0
  for i, j in zip(df.loc[:, 'Code'], df.loc[:, 'Amount']):
    for k in svdf[f'House Number  {house_number}']:
      if str(i).lower() in str(k).lower():
        svdf.loc[svdf[f'House Number  {house_number}'] == str(k), 'This Period'] += j
        print(svdf.loc[svdf[f'House Number  {house_number}'] == str(k), :])
  return svdf


def update_schedule(svdf, house_number):
  for i in svdf[f'House Number {house_number}']:
    total = svdf.loc[svdf[f'House Number  {house_number}'] == str(i), 'This Period'] + svdf.loc[svdf[f'House Number  {house_number}'] == str(i), 'Previous Period']
    svdf.loc[svdf[f'House Number  {house_number}'] == str(i), 'Total Billed'] = round(total, 2)
    budget = svdf.loc[svdf[f'House Number  {house_number}'] == str(i), 'Scheduled Value']
    percent = round(100 * (total / budget))
    svdf.loc[svdf[f'House Number {house_number}'] == str(i), 'Percent Billed'] = percent
    svdf.loc[svdf[f'House Number {house_number}'] == str(i), 'Balance to Finish'] = round(budget - total, 2)
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


def set_budget(svdf, house_number):
  for i in svdf[f'House Number {house_number}']:
    if i == 'Total':
      continue
    budget = input(f'What is the budget for {i}? ')
    budget = float(budget)
    svdf.loc[svdf[f'House Number {house_number}'] == str(i), 'Scheduled Value'] = budget
  svdf = get_totals(svdf, True)
  return svdf


if __name__ == '__main__':
  main()
