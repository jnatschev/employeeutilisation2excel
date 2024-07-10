"""

"""
from time import sleep
from employee_utilisation import FortnightlyEmployeeUtilisation
from openpyxl.utils.dataframe import dataframe_to_rows
import pathlib
import openpyxl


def employees():
    employee_file = pathlib.Path('../employees.txt')
    with open(employee_file, mode='r') as fd:
        for _employee in fd.read().split('\n'):
            yield _employee


if __name__ == "__main__":
    for employee in employees():
        print('Employee {} start'.format(employee))
        feu = FortnightlyEmployeeUtilisation(name=employee)
        employee_excel_file = pathlib.Path(
            '{} - Fortnightly - {} - Utilisation.xlsx'.format(feu.financial_year, feu.name)
        )
        if not employee_excel_file.exists():
            feu.as_df.to_excel(excel_writer=employee_excel_file, index=False)
        else:
            workbook = openpyxl.load_workbook(employee_excel_file)
            worksheet = workbook.active

            for row in dataframe_to_rows(df=feu.as_df, header=False, index=False):
                worksheet.append(row)

            workbook.save(employee_excel_file)
            workbook.close()
        print('Employee {} end'.format(employee))
        print()
        sleep(5.0)
