import pandas as pd
import re


def clean_date(deadlines_lst: list):
    """
    :param deadlines_lst:
    :return:
    """
    lst = []
    for deadline in deadlines_lst:
        deadline = deadline.strip()
        date_pattern = r'[A-Za-z]+\s*\d{1,2}'
        match = re.search(date_pattern, deadline, re.IGNORECASE)
        if match:
            lst.append(match.group())
        else:
            lst.append('Ongoing')
    return lst


def clean_numbers(column_lst: list):
    """
    :param column_lst:
    :return:
    """
    lst = []
    for text in column_lst:
        text = text.lstrip(".0123456789 ")
        lst.append(text)
    return lst


if __name__ == '__main__':
    df = pd.read_excel('okr_table_data.xlsx')
    df['Deadline'] = pd.Series(clean_date(df['Deadline']))
    df.to_excel('okr_data_dates_only_clean_.xlsx', index=False)

    df['OKRs'] = pd.Series(clean_numbers(df['OKRs']))
    df['Projects'] = pd.Series(clean_numbers(df['Projects']))

    df.to_excel('okr_data_all_clean.xlsx', index=False)
