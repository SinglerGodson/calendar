import os

import pandas as pd

from src.fill_ppt import fill_ppt
# print('test')

# os.listdir('./')

user_input_path = input("请输入文件路径: ")
# user_input_path = '/Users/mac/Documents/学校行政/校行事历/2025-2026学年第一学期第十一周、第十二周行事历.xlsx'
df = pd.read_excel(user_input_path, 0)


columns = df.columns

week_col = columns[0]
date_col = columns[1]
weekday_col = columns[3]




def content(content):
    # 跳过空值
    if pd.isna(content):
        return ''

    # 转换为字符串
    c = str(content).strip()

    # 跳过空字符串
    if not c or c == '/':
        return ''

    return c

dept = ''
week = {}
date = {}
weekday = {}
calendars = []

for i, column in enumerate(columns):

    column_data = df[column]

    # for j, cell_content in enumerate(column_data):
    #     c = content(cell_content)
    #     print(f'row: {i }, col: {j }, content: {c }')

    # 周次
    if i == 0:
        for j, cell_content in enumerate(column_data):
            c = content(cell_content)

            if c != '':
                week[j] = c

    # 日期
    if i == 1:
        # print(f'日期： {column}')
        for d, cell_content in enumerate(column_data):
            date[d] = content(cell_content)

    # 星期
    if i == 2:
        for w, cell_content in enumerate(column_data):
            weekday[w] =  content(cell_content)


    if i in (3, 6, 9, 12, 15, 18, 21, 24, 27, 30):

        # column_data = df[column]
        for j, cell_content in enumerate(column_data):

            calendar = {}

            cell_content = content(cell_content)
            if j == 1:
                dept = cell_content
                # continue

            if j < 2:
                continue

            if j >= 16:
                break

            for k, v in week.items():
                if j >= k:
                    # print(f'week: {v }')
                    calendar['week'] = v
                    # break

            for k, v in weekday.items():
                if j >= k:
                    # print(f'weekday：{weekday[j ] }')
                    calendar['weekday'] = weekday[j]
                    # break

            if j in date.keys():
                calendar['date'] = date[j]

            # print(f'date：{v }')
            calendar['dept'] = dept

            # print(f'{cell_content }\n')

            calendar['content'] = cell_content

            calendars.append(calendar)


dirname, basename = os.path.split(user_input_path)
basename_no_ext, ext = os.path.splitext(basename)

# 替换扩展名
# 组合新的完整文件名
new_filename_with_path = os.path.join(dirname, basename_no_ext + ".pptx")


fill_ppt(calendars, new_filename_with_path)
# print(calendars)



