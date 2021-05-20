#個人の投票からポイント合計してチームの順位を求めるプログラム

import openpyxl

wb = openpyxl.load_workbook('/Users/sakaiyukito/Downloads/python_study/オンライン合宿順位投票（回答）.xlsx')
ws = wb['Sheet2']

team_list = ['1-a', '1-b', '1-c', '1-d', '2-a', '2-b', '2-c', '3-a', '3-b', '3-c', '4-a', '4-b', '4-c', '4-d']
sum = 0
sum_list = []
for col in ws.iter_cols(min_row=2, min_col=4):
    for cell in col:
        if cell.value == '1位':
            sum += 14
        elif cell.value == '2位':
            sum += 13
        elif cell.value == '3位':
            sum += 12
        elif cell.value == '4位':
            sum += 11
        elif cell.value == '5位':
            sum += 10
        elif cell.value == '6位':
            sum += 9
        elif cell.value == '7位':
            sum += 8
        elif cell.value == '8位':
            sum += 7
        elif cell.value == '9位':
            sum += 6
        elif cell.value == '10位':
            sum += 5
        elif cell.value == '11位':
            sum += 4
        elif cell.value == '12位':
            sum += 3
        elif cell.value == '13位':
            sum += 2
        elif cell.value == '14位':
            sum += 1

    sum_list.append(sum)
    sum = 0
    
print(sum_list)

dict = dict(zip(team_list, sum_list))

dict = sorted(dict.items(), key=lambda x:x[1], reverse=True)
for i, j in enumerate(dict):
    print(i+1,"位：" , j)
