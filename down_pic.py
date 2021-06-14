

from urllib.request import urlretrieve
from PIL import Image
import openpyxl
import ssl
ssl._create_default_https_context = ssl._create_unverified_context


file='自如租房20210614230548.xlsx'
def read():
    workbook = openpyxl.load_workbook(file)
    booksheet = workbook.active


    total =[]
    # 迭代所有的行
    for i, row in enumerate(booksheet.rows):
        if i ==0:
            continue
        line = [col.value for col in row]
        print(line[12])
        urlretrieve(line[12],'./pic/room' + str(i) + ".webp")
        urlretrieve(line[13],'./pic/house' + str(i) + ".webp")

        im = Image.open('./pic/room' + str(i) + ".webp")  # 读入文件
        im.save('./pic/room' + str(i) + ".jpg")  # 解码保存

        im = Image.open('./pic/house' + str(i) + ".webp")  # 读入文件
        im.save('./pic/house' + str(i) + ".jpg")  # 解码保存
        #12 13
        line[12] = 'room' + str(i) + ".jpg"
        line[13] = 'house' + str(i) + ".jpg"

        total.append(line)

    write(total)


def write(total):
    print('正在保存')

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet['A1'] = "题名"
    sheet['B1'] = "价格"
    sheet['C1'] = "第几室"
    sheet['D1'] = "面积"
    sheet['E1'] = "小区"
    sheet['F1'] = "朝向"
    sheet['G1'] = "户型"
    sheet['H1'] = "楼层"
    sheet['I1'] = "电梯"
    sheet['J1'] = "可入住日期"
    sheet['K1'] = "签约时长"
    sheet['L1'] = "室友信息"
    sheet['M1'] = "屋内pic"
    sheet['N1'] = "户型pic"
    sheet['O1'] = "阳台"
    sheet['P1'] = "独卫"
    sheet['Q1'] = "链接"

    for i in total:
        sheet.append(i)

    workbook.save("jpg版本.xlsx")
    print('保存成功')


if __name__ == '__main__':
    read()