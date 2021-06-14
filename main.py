
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains

import time
import re
import openpyxl

center_keyword = '苏豪名厦'
map_link = 'https://sz.ziroom.com/map/'
chromedriver_path = './chromedriver'
search_link = 'https://sz.ziroom.com/z/?qwd=%E5%9D%82%E7%94%B0'

def show_map():
    driver.get(map_link)
    driver.maximize_window()
    elm = driver.find_element_by_xpath('//*[@id="J_MyKeywords"]')
    elm.click()
    elm.send_keys(center_keyword)
    time.sleep(1)
    elm.send_keys(Keys.ENTER)




def get_community():
    time.sleep(2)
    elms = driver.find_elements_by_class_name('houseMarkers')


    file = open('community.txt','w',encoding='utf-8')

    total_num = 0
    for elm in elms:
        pattern = re.compile(r'<[^>]+>', re.S)
        text = pattern.sub(' ', elm.get_attribute("innerHTML"))
        textlist=text.split()
        name = textlist[0]
        num = textlist[-1].replace('间','')
        file.write(name+' '+num+'\n')
        total_num += int(num)
    file.write(str(total_num))
    file.close()

def dragger_all_house():

    driver.find_element_by_id('mCSB_1_scrollbar_vertical').click()

    action_chains = ActionChains(driver)
    for i in range(40):
        action_chains.key_down(Keys.DOWN).perform()
        time.sleep(2)
        action_chains.key_up(Keys.DOWN).perform()
        time.sleep(1)

        scrollbar_height = driver.find_element_by_id('mCSB_1_scrollbar_vertical').size['height']
        dragger = driver.find_element_by_id('mCSB_1_dragger_vertical')
        min_height = dragger.value_of_css_property('min-height').replace('px','')
        top = dragger.value_of_css_property('top').replace('px','')
        if int(scrollbar_height) - int(min_height) == int(top):
            break

def get_all_house():
    file = open('house.txt', 'w', encoding='utf-8')

    ul = driver.find_element_by_id('J_houseList')

    house_list = []
    for li in ul.find_elements_by_tag_name('li'):
        house_link = li.find_element_by_tag_name('a').get_attribute('href')
        price = int(li.find_element_by_class_name('org').get_attribute("innerHTML").replace('¥',''))
        if price>3000 or price<1500:
            continue
        file.write(str(price) + ' ' +house_link+'\n')
        house_list.append([house_link,price])
    file.close()

    print('共有'+str(len(house_list))+'间房')

    write(get_details(house_list))

def get_all_house_by_file():
    file = open('house.txt', 'r', encoding='utf-8')

    house_list = []

    for line in file.readlines():
        price,link = line.strip().split(' ')
        house_list.append([link,price])

    file.close()

    print('共有'+str(len(house_list))+'间房')

    write(get_details(house_list))


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
        sheet.append([i['题名'], i['价格'],i['第几室'],i['面积'],i['小区'],i['朝向'],i['户型'],i['楼层'],i['电梯'],i['可入住日期'],
                      i['签约时长'],i['室友信息'],i['屋内pic'],i['户型pic'],i['阳台'],i['独卫'],i['链接']])

    workbook.save("自如租房"+ time.strftime("%Y%m%d%H%M%S", time.localtime()) +".xlsx")
    print('保存成功')


def get_details(house_list):
    result =[]
    for i in house_list:
        time.sleep(5)
        try:
            info = get_detail(i[0])
            print(i[0],i[1])
            info['价格']=i[1]
            info['链接'] = i[0]
            result.append(info)
        except:
            print("报错："+i[0])
            break
    return result


def get_detail(link):

    res ={}
    # 可入住日期 签约时长 室友信息
    driver.get(link)

    title = driver.find_element_by_xpath('/html/body/section/aside/h1').text #题名
    res['题名']=title

    res['第几室']= title.split('-')[-1]

    res['面积'] = driver.find_element_by_xpath('/html/body/section/aside/div[3]/div[1]/dl[1]/dd').text  #使用面积

    res['小区'] = driver.find_element_by_xpath('//*[@id="villageinfo"]/div/div/div/h3').text    #小区

    res['朝向'] = driver.find_element_by_xpath('/html/body/section/aside/div[3]/div[1]/dl[2]/dd').text #朝向

    res['户型'] = driver.find_element_by_xpath('/html/body/section/aside/div[3]/div[1]/dl[3]/dd').text #户型

    res['楼层'] = driver.find_element_by_xpath('/html/body/section/aside/div[3]/ul/li[2]/span[2]').text  #楼层

    res['电梯'] = driver.find_element_by_xpath('/html/body/section/aside/div[3]/ul/li[3]/span[2]').text  #电梯


    if len(driver.find_element_by_xpath('//*[@id="live-tempbox"]/ul').find_elements_by_tag_name('li'))==2:
        res['可入住日期'] =''
        res['签约时长'] = driver.find_element_by_xpath('//*[@id="live-tempbox"]/ul/li[1]/span[2]').text
    else:
        res['可入住日期'] = driver.find_element_by_xpath('//*[@id="live-tempbox"]/ul/li[1]/span[2]').text  #可入住日期
        res['签约时长'] = driver.find_element_by_xpath('//*[@id="live-tempbox"]/ul/li[2]/span[2]').text  #签约时长


    people_info =''
    for li in driver.find_element_by_xpath('//*[@id="meetinfo"]/div/ul').find_elements_by_tag_name('li'):
        text1 = li.get_attribute("innerHTML")
        if '入住' in text1:
            if '男' in text1:
                people_info+='男 '
            elif '女' in text1:
                people_info += '女 '
            else:
                people_info += '不知 '
        else:
            people_info += '无 '

    res['室友信息']= people_info

    pics = driver.find_element_by_xpath('//*[@id="Z_swiper_box"]/div[2]/div/ul').find_elements_by_tag_name('li')  #屋内pic

    huxing_pic =''
    for li in pics:
        if li.get_attribute('data-t')=='户型':
            huxing_pic = li.find_element_by_tag_name('img').get_attribute('src')
            break
    res['户型pic'] = huxing_pic

    house_pic=''
    for li in pics:
        if li.get_attribute('data-t') == '图片':
            house_pic = li.find_element_by_tag_name('img').get_attribute('src')
            break
    res['屋内pic']= house_pic

    tags =driver.find_element_by_xpath('/html/body/section/aside/div[2]').get_attribute("innerHTML")
    yangtai = '有' if '阳台' in tags else '无'   #阳台
    res['阳台']= yangtai

    weishengjian = '有' if '独卫' in tags else '无'   #阳台  #卫生间
    res['独卫']= weishengjian

    return res





if __name__ == '__main__':
    driver = webdriver.Chrome(executable_path=chromedriver_path)

    # show_map()
    # get_community()
    # dragger_all_house()
    # get_all_house()

    get_all_house_by_file()

    driver.close()
