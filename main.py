from utils import createExcel, findRoots, loadExcel

#Run python main.py to try code on the sample below
filename = 'quadratic_data.xlsx'
sheet1 = 'Sheet1'
sheet2 = 'Sheet2'

createExcel(filename, sheet1, sheet2)
roots_list = findRoots(filename, sheet1, sheet2)
loadExcel(filename, sheet1, sheet2, roots_list)
