from Excel import OpenSheet
import math

wb = OpenSheet.wb
ws = OpenSheet.ws

max_row = 2002
max_col = 7

for col in ws.iter_cols(min_col=3, max_col=7):
	position = []
	dataSum = 0  # 数字维度的和
	dataAvg = 0  # 数字维度的平均值
	for cell in col:
		val = cell.value
		# 终止条件
		if val == None:
			break
		# 跳过表头
		if isinstance(val, str) and val != '?':
			continue
		if val == '?':
			position.append(cell.column)
		else:
			dataSum += val

	dataAvg = math.floor(dataSum / (max_row - 2 - len(position)))

	for row in position:
		ws.cell(row=row, column=col[0].column).value = dataAvg

wb.save('数据清理后的数据集.xlsx')
wb.close()