#明列番号順に2乗誤差を求めるプログラム

import openpyxl

wb = openpyxl.load_workbook('/Users/sakaiyukito/Downloads/python_study/オンライン合宿順位投票（回答）.xlsx')
ws = wb['Sheet2']

team_list = ['1-a', '1-b', '1-c', '1-d', '2-a', '2-b', '2-c', '3-a', '3-b', '3-c', '4-a', '4-b', '4-c', '4-d']
rank_all = [14, 11, 12, 2, 7, 1, 10, 4, 5, 9, 6, 8, 3, 13]
rank = []
error_list = []
print('-------------------------------------')

for row in ws.iter_rows(min_row=2, min_col=4, max_row=56, max_col=17):
    rank_list = []
    for cell in row:
        if cell.value == '1位':
            rank_list.append(1)
        elif cell.value == '2位':
            rank_list.append(2)
        elif cell.value == '3位':
            rank_list.append(3)
        elif cell.value == '4位':
            rank_list.append(4)
        elif cell.value == '5位':
            rank_list.append(5)
        elif cell.value == '6位':
            rank_list.append(6)
        elif cell.value == '7位':
            rank_list.append(7)
        elif cell.value == '8位':
            rank_list.append(8)
        elif cell.value == '9位':
            rank_list.append(9)
        elif cell.value == '10位':
            rank_list.append(10)
        elif cell.value == '11位':
            rank_list.append(11)
        elif cell.value == '12位':
            rank_list.append(12)
        elif cell.value == '13位':
            rank_list.append(13)
        elif cell.value == '14位':
            rank_list.append(14)

    rank.append(rank_list)
#print(rank)

for i in range(55):
    error = 0
    for j in range(14):
        error  += (rank_all[j] - rank[i][j]) ** 2
    error_list.append(error)

print(error_list)

for i, j in enumerate(error_list):
    print('明列番号', i+1, '： 誤差', j)