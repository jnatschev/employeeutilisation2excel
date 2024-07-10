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
    if 1 <= today.day <= 14 and today.month == 7:
        year = today.year
        return AuVicFinancialYear(year=year)
    elif today.month >= 7:
        return AuVicFinancialYear(year=today.year + 1)
    else:
        return AuVicFinancialYear(year=today.year)


def return_sql_query_start_date():
    if 1 <= today.day <= 14:
        year = (today - DateOffset(months=1)).year
        month = (today - DateOffset(months=1)).month
        sql_query_start_date = Timestamp('{}-{}-15'.format(year, month)).normalize()
    else:
        month = financial_year.today.month
        mask = (financial_year.calendar.date.dt.month == month)
        sql_query_start_date = Timestamp(
            financial_year.calendar.date[mask].min().date().isoformat()
        )
    return sql_query_start_date


def return_sql_query_end_date():
    if 1 <= today.day <= 14:
        month = (today - DateOffset(months=1)).month
        if today.month == 7:
            month_last_day = financial_year.calendar.date.max().date().isoformat()
        else:
            mask = financial_year.calendar.date.dt.month == month
            month_last_day = financial_year.calendar.date[mask].max().date().isoformat()
        sql_query_end_date = Timestamp('{}'.format(month_last_day))
    else:
        year = financial_year.today.year
        month = financial_year.today.month
        sql_query_end_date = Timestamp('{}-{}-14'.format(year, month))
    return sql_query_end_date


if __name__ == "__main__":
    today = Timestamp('today').normalize()
    financial_year = return_financial_year()
    start_date = return_sql_query_start_date()
    end_date = return_sql_query_end_date()
    for employee in employees():
        feu = EmployeeUtilisation(
            end_date=end_date,
            financial_year=financial_year,
            name=employee,
            start_date=start_date,
            today=today
        )
        employee_excel_file = pathlib.Path(
            '../{} - Fortnightly - {} - Utilisation.xlsx'.format(feu.financial_year.year, feu.name)
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
