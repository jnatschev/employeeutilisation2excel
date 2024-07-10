"""

"""
from employee_utilisation import EmployeeUtilisation
from financial_year import AuVicFinancialYear
from openpyxl.utils.dataframe import dataframe_to_rows
from pandas import DateOffset, Timestamp
import pathlib
import openpyxl


def employees():
    employee_file = pathlib.Path('../employees.txt')
    with open(employee_file, mode='r') as fd:
        for _employee in fd.read().split('\n'):
            yield _employee


def return_financial_year():
    if today.month == 7:
        year = today.year
        return AuVicFinancialYear(year=year)
    elif 8 <= today.month <= 12:
        return AuVicFinancialYear(year=today.year + 1)
    else:
        return AuVicFinancialYear(year=today.year)


def return_sql_query_start_date():
    month = (today - DateOffset(months=1)).month
    mask = (financial_year.calendar.date.dt.month == month)
    month_df = financial_year.calendar.date[mask]
    sql_query_start_date = month_df.min().date()
    return sql_query_start_date


def return_sql_query_end_date():
    month = (today - DateOffset(months=1)).month
    mask = (financial_year.calendar.date.dt.month == month)
    month_df = financial_year.calendar.date[mask]
    sql_query_end_date = month_df.max().date()
    return sql_query_end_date


if __name__ == "__main__":
    today = Timestamp('today').normalize()
    financial_year = return_financial_year()
    start_date = return_sql_query_start_date()
    end_date = return_sql_query_end_date()
    for employee in ('John Natschev'):
    # for employee in employees():
        feu = EmployeeUtilisation(
            end_date=end_date,
            financial_year=financial_year,
            name=employee,
            start_date=start_date,
            today=today
        )
        employee_excel_file = pathlib.Path(
            '../{} - Monthly - {} - Utilisation.xlsx'.format(feu.financial_year.year, feu.name)
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
