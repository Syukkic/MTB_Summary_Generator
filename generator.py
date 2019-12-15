import numpy as np
import pandas as pd
from dateutil.parser import parse

class Generator:
    def __init__(self):
        self.filename = None
    
    def read_file(self, filename):
        try:
            data = pd.read_excel(str(filename), sheet_name=1, skiprows=1)
            data.index += 1
            return data
        except:
            return "Your file went wrong!"

    def get_product_name(self, data):
        product = data['Product'].unique()[0]
        return product

    def get_date(self, data):
        months = {
            1:'Jan.', 2:'Feb.', 3:'Mar.', 4:'Apr.', 5:'May', 6:'Jun.',
            7:'Jul.', 8:'Aug.', 9:'Sep.', 10:'Oct.', 11:'Nov.', 12:'Dec.'
        }

        date = str(data['Month'].unique()[0])
        year = parse(date, fuzzy = True).year
        month = months[parse(date, fuzzy = True).month]

        return year, month

    def extract_specs(self, data):
        for index, row in data.iterrows():
            if row['Specification'][-5:] == 'Tech.':
                data.at[index, 'New_Spec'] = 'Technical'
            else:
                data.at[index, 'New_Spec'] = row['Specification']
        
        unique_spec = sorted(list(data['New_Spec'].unique()), reverse = True)
        print('Specification are ', unique_spec)

        return unique_spec

    def pivot_table(self, data, unique_spec):

        tables = []

        for n in range(len(unique_spec)):
            #  Generate destination pivot table
            pt_destination = pd.pivot_table(data[data['New_Spec'] == unique_spec[n]], index = ['Destination'], columns = ['Buyer'], values = ['Quantity'], aggfunc = [np.sum], fill_value = 0, margins = True, margins_name = 'Total')
            pt_destination = pt_destination.sort_values(('sum', 'Quantity', 'Total'), ascending = False)

            # ranking the buyer
            pt_destination.T.sort_values(('Buyer'), ascending = True).T
            pt_destination = pt_destination

            # Move the top total to the bottom
            total = pt_destination.iloc[[0], :]
            pt_destination = pt_destination.drop(index = 'Total')
            pt_destination = pt_destination.append(total)
            tables.append(pt_destination)

            #  Generate exporter pivot table
            pt_exporter = pd.pivot_table(data[data['New_Spec'] == unique_spec[n]], index = ['Company'], columns = ['Buyer'], values = ['Quantity'], aggfunc = [np.sum], fill_value = 0, margins = True, margins_name = 'Total')
            pt_exporter = pt_exporter.sort_values(('sum', 'Quantity', 'Total'), ascending = False)
            pt_exporter.index.name = 'Exporter'

            # ranking the buyer
            pt_exporter.T.sort_values(('Buyer'), ascending = True).T
            pt_exporter = pt_exporter

            # Move the top total to the bottom
            total = pt_exporter.iloc[[0], :]
            pt_exporter = pt_exporter.drop(index = 'Total')
            pt_exporter = pt_exporter.append(total)
            tables.append(pt_exporter)

        return tables

    def save_to_excel(self, product, year, month, tables, unique_spec):
        writer = pd.ExcelWriter("result.xlsx", engine = 'xlsxwriter')
        row = 2
        table_number = 0
        table_names = ['Exporters and buyers', 'Destinations and buyers']
        table_spec = unique_spec
        table_spec.extend(unique_spec)
        table_spec = sorted(table_spec, reverse = True)
        
        for table in tables:
            table.to_excel(writer, sheet_name = 'Summary', startrow = row, startcol = 0)
            row = row + len(table.index) + 7
            table_number += 1

            print('Table {} {} of {} {} in {} {} (kg)'.format(table_number, table_names[table_number % 2], product, table_spec[table_number-1], month, year))
            print('')

        writer.save()