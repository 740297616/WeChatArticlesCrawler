import requests
from bs4 import BeautifulSoup
import time
import random
from openpyxl import Workbook
import tkinter as tk
from tkinter import scrolledtext
import sys
import re

# 23.7.27 增加小尾巴处理函数
def footing_rebuild(text):

    def get_first_element_or_empty(my_list):
        return my_list[0] if my_list else ''

    # text = "撰文 WRITER苏颖琛、石小曼采写 WRITER陈琳涵摄影 PHOTOGRAPHER刘曦茗、倪浩源、易相宇、黄琪琛视频 VIDEO祁毅恒排版 EDITOR陈晓、萧颖琦设计 DESIGNER卢胤瑜、林艾莹、刘国林蔡婕、谭俊熙手绘 PAINTER钟小桦、蓝文懿技术支持 TECH SUPPORTER黄靖钧、纪元、曾健欢"

    # 替换“字符串”为“ 字符串”
    text = text.replace(" PLANNER", ": ")
    text = text.replace(" WRITER", ": ")
    text = text.replace(" PHOTOGRAPHER", ": ")
    text = text.replace(" VIDEO", ": ")
    text = text.replace(" EDITOR", ": ")
    text = text.replace(" SVG DESIGNER", ": ")
    text = text.replace(" DESIGNER", ": ")
    text = text.replace(" GIF", ": ")
    text = text.replace(" PHOTO ORGANIZER", ": ")
    text = text.replace(" PAINTER", ": ")
    text = text.replace(" SOURCE", ": ")
    text = text.replace(" TECH SUPPORTER", ": ")

    # 在"采写"前面添加空格
    text = text.replace("采写", " 采写")
    # 在"摄影"前面添加空格
    text = text.replace("摄影", " 摄影")
    # 在"视频"前面添加空格
    text = text.replace("视频", " 视频")
    # 在"排版"前面添加空格
    text = text.replace("排版", " 排版")
    # 在"设计"前面添加空格
    text = text.replace("设计", " 设计")
    # 在"手绘"前面添加空格
    text = text.replace("手绘", " 手绘")
    # 在"技术支持"前面添加空格
    text = text.replace("技术支持", " 技术支持")
    # 在"策划"前面添加空格
    text = text.replace("策划", " 策划")
    # 在"设计"前面添加空格
    text = text.replace("设计", " 设计")
    # 在"交互"前面添加空格
    text = text.replace("交互", " 交互")
    # 在"GIF制作"前面添加空格
    text = text.replace("GIF制作", " GIF制作")
    # 在"图片整理"前面添加空格
    text = text.replace("图片整理", " 图片整理")
    # 在"图片来源"前面添加空格
    text = text.replace("图片来源", " 图片来源")
    # 在"来源"前面添加空格
    text = text.replace("来源", " 来源")

    tasks = ["撰文", "采写", "摄影", "视频", "排版", "设计", "手绘", "技术支持", "策划", "交互", "GIF制作",
             "图片整理", "图片来源", "来源"]
    name_dict = {task: [] for task in tasks}

    for task in tasks:
        pattern = re.compile(rf"{task}: ([\u4e00-\u9fa5、]+)")
        matches = pattern.findall(text)
        name_dict[task].extend(matches)

    # 分别保存到不同的变量
    authors = name_dict["撰文"]
    writers = name_dict["采写"]
    photographers = name_dict["摄影"]
    videographers = name_dict["视频"]
    typographers = name_dict["排版"]
    designers = name_dict["设计"]
    illustrators = name_dict["手绘"]
    technical_support = name_dict["技术支持"]
    planner = name_dict["策划"]
    svg_designer = name_dict["交互"]
    gif_maker = name_dict["GIF制作"]
    photo_organizer = name_dict["图片整理"]
    photo_source = name_dict["图片来源"]
    source = name_dict["来源"]

    # 去除变量中的方括号和单引号，否则写入不了excel
    authors = get_first_element_or_empty(authors).replace('[', '').replace(']', '').replace("'", '')
    writers = get_first_element_or_empty(writers).replace('[', '').replace(']', '').replace("'", '')
    photographers = get_first_element_or_empty(photographers).replace('[', '').replace(']', '').replace("'", '')
    videographers = get_first_element_or_empty(videographers).replace('[', '').replace(']', '').replace("'", '')
    typographers = get_first_element_or_empty(typographers).replace('[', '').replace(']', '').replace("'", '')
    designers = get_first_element_or_empty(designers).replace('[', '').replace(']', '').replace("'", '')
    illustrators = get_first_element_or_empty(illustrators).replace('[', '').replace(']', '').replace("'", '')
    technical_support = get_first_element_or_empty(technical_support).replace('[', '').replace(']', '').replace("'", '')
    planner = get_first_element_or_empty(planner).replace('[', '').replace(']', '').replace("'", '')
    svg_designer = get_first_element_or_empty(svg_designer).replace('[', '').replace(']', '').replace("'", '')
    gif_maker = get_first_element_or_empty(gif_maker).replace('[', '').replace(']', '').replace("'", '')
    photo_organizer = get_first_element_or_empty(photo_organizer).replace('[', '').replace(']', '').replace("'", '')
    photo_source = get_first_element_or_empty(photo_source).replace('[', '').replace(']', '').replace("'", '')
    source = get_first_element_or_empty(source).replace('[', '').replace(']', '').replace("'", '')

    return authors, writers, photographers, videographers, typographers, designers, illustrators, technical_support \
        , planner, svg_designer, gif_maker, photo_organizer, photo_source, source

# Step2 进一步获取文章点赞、阅读、在看数和小尾巴
def getMoreInfo(link):
    # 获得mid,_biz,idx,sn 这几个在link中的信息
    mid = link.split("&")[1].split("=")[1]
    idx = link.split("&")[2].split("=")[1]
    sn = link.split("&")[3].split("=")[1]
    _biz = link.split("&")[0].split("_biz=")[1]

    pass_ticket = WX_Pass_ticket
    appmsg_token = WX_Appmsg_Token
    url = "http://mp.weixin.qq.com/mp/getappmsgext" # 获取详情页的网址
    phoneCookie = WX_Cookie
    headers = {
        "Cookie": phoneCookie,
        "User-Agent": WX_User_Agent
    }
    data = {
        "is_only_read": "1",
        "is_temp_url": "0",
        "appmsg_type": "9",
        'reward_uin_count': '0'
    }
    params = {
        "__biz": _biz,
        "mid": mid,
        "sn": sn,
        "idx": idx,
        "key": WX_Key,
        "pass_ticket": pass_ticket,
        "appmsg_token": appmsg_token,
        "uin": "MjgwMjI0MTMyNQ==",
        "wxtoken": "777",
    }

    response = requests.get(link, verify=False)
    soup = BeautifulSoup(response.content, 'html.parser')

    # 获取整个页面的文本
    page_text = soup.get_text()

    # 找到"撰文"和"版权"之间的文本
    start_index = page_text.find("撰文")
    end_index = page_text.find("责编")
    try:
        HtmlFoot = page_text[start_index:end_index]
        print(HtmlFoot)
    except:
        HtmlFoot = 0

    authors, writers, photographers, videographers, typographers, designers, illustrators, technical_support , planner\
        ,svg_designer, gif_maker, photo_organizer, photo_source, source = footing_rebuild(HtmlFoot)

    requests.packages.urllib3.disable_warnings()
    # 使用post方法进行提交，这里返回了一个json，里面是单个文章的数据
    content = requests.post(url, headers=headers, data=data, params=params).json()

    # 在上面返回的json中获取并输出阅读数点赞数在读数
    try:
        readNum = content["appmsgstat"]["read_num"]
        print("阅读数:" + str(readNum))
    except:
        readNum = 0
    try:
        likeNum = content["appmsgstat"]["like_num"]
        print("点赞数:" + str(likeNum))
    except:
        likeNum = 0
    try:
        old_like_num = content["appmsgstat"]["old_like_num"]
        print("在读数:" + str(old_like_num))
    except:
        old_like_num = 0

    time.sleep(3) # 歇3s，防止被封
    return readNum, likeNum, old_like_num, authors, writers, photographers, videographers, typographers, \
        designers, illustrators, technical_support , planner, svg_designer, gif_maker, photo_organizer, photo_source, \
        source


# Step1 批量获取文章url,title,create_time
def getAllInfo():
    url = "https://mp.weixin.qq.com/cgi-bin/appmsg"
    Cookie = platform_Cookie
    headers = {
        "Cookie": Cookie,
        "User-Agent": platform_UserAgent
    }
    token = platform_Token  # 公众号
    fakeid = platform_Fakeid  # 公众号对应的id
    type = '9'
    data1 = {
        "token": token,
        "lang": "zh_CN",
        "f": "json",
        "ajax": "1",
        "action": "list_ex",
        "begin": "0",
        "count": "4",
        "query": "",
        "fakeid": fakeid,
        "type": type,
    }

    # 拿一页，存一页
    messageAllInfo = []
    # begin 从0开始
    for i in range(33):  # 设置爬虫页码
        begin = i * 4
        data1["begin"] = begin
        requests.packages.urllib3.disable_warnings()
        content_json = requests.get(url, headers=headers, params=data1, verify=False).json()
        time.sleep(random.randint(1, 10))
        if "app_msg_list" in content_json:
            for item in content_json["app_msg_list"]:
                timestamp = item['create_time']
                time_local = time.localtime(timestamp)
                spider_url = item['link']
                readNum, likeNum, old_like_num, authors, writers, photographers, videographers, typographers,\
                    designers, illustrators, technical_support , planner, svg_designer, gif_maker, photo_organizer, \
                    photo_source, source = getMoreInfo(spider_url)
                info = {
                    "title": item['title'],
                    "createTime": time.strftime("%Y-%m-%d %H:%M:%S",time_local),
                    "url": item['link'],
                    "readNum": readNum,
                    "likeNum": likeNum,
                    "old_like_num": old_like_num,
                    "authors": authors,
                    "writers": writers,
                    "photographers": photographers,
                    "videographers": videographers,
                    "typographers": typographers,
                    "designers": designers,
                    "illustrators": illustrators,
                    "technical_support": technical_support,
                    "planner": planner,
                    "svg_designer": svg_designer,
                    "gif_maker": gif_maker,
                    "photo_organizer": photo_organizer,
                    "photo_source": photo_source,
                    "source": source
                }
                messageAllInfo.append(info)
    return messageAllInfo

def main():
    f = Workbook()  # 创建一个workbook 设置编码
    sheet = f.active  # 创建sheet表单
    # 写入表头
    sheet.cell(row=1, column=1).value = 'title(推文标题)'  # 第一行第一列
    sheet.cell(row=1, column=2).value = 'creatTime(发布时间)'
    sheet.cell(row=1, column=3).value = 'url(推文链接)'
    sheet.cell(row=1, column=4).value = 'readNum(阅读数)'
    sheet.cell(row=1, column=5).value = 'likeNum(点赞数)'
    sheet.cell(row=1, column=6).value = 'old_like_num(在看数)'
    sheet.cell(row=1, column=7).value = '策划'
    sheet.cell(row=1, column=8).value = '撰文'
    sheet.cell(row=1, column=9).value = '采写'
    sheet.cell(row=1, column=10).value = '排版'
    sheet.cell(row=1, column=11).value = '摄影'
    sheet.cell(row=1, column=12).value = '手绘'
    sheet.cell(row=1, column=13).value = '设计'
    sheet.cell(row=1, column=14).value = '交互'
    sheet.cell(row=1, column=15).value = '视频'
    sheet.cell(row=1, column=16).value = 'GIF制作'
    sheet.cell(row=1, column=17).value = '技术支持'
    sheet.cell(row=1, column=18).value = '图片整理'
    sheet.cell(row=1, column=19).value = '图片来源'
    sheet.cell(row=1, column=20).value = '来源'
    messageAllInfo = getAllInfo()   # 获取信息
    # print(messageAllInfo)
    print(len(messageAllInfo))  # 输出列表长度
    # 写内容
    for i in range(1, len(messageAllInfo)+1):
        sheet.cell(row=i + 1, column=1).value = messageAllInfo[i - 1]['title']
        sheet.cell(row=i + 1, column=2).value = messageAllInfo[i - 1]['createTime']
        sheet.cell(row=i + 1, column=3).value = messageAllInfo[i - 1]['url']
        sheet.cell(row=i + 1, column=4).value = messageAllInfo[i - 1]['readNum']
        sheet.cell(row=i + 1, column=5).value = messageAllInfo[i - 1]['likeNum']
        sheet.cell(row=i + 1, column=6).value = messageAllInfo[i - 1]['old_like_num']
        sheet.cell(row=i + 1, column=7).value = messageAllInfo[i - 1]['planner']
        sheet.cell(row=i + 1, column=8).value = messageAllInfo[i - 1]['authors']
        sheet.cell(row=i + 1, column=9).value = messageAllInfo[i - 1]['writers']
        sheet.cell(row=i + 1, column=10).value = messageAllInfo[i - 1]['typographers']
        sheet.cell(row=i + 1, column=11).value = messageAllInfo[i - 1]['photographers']
        sheet.cell(row=i + 1, column=12).value = messageAllInfo[i - 1]['illustrators']
        sheet.cell(row=i + 1, column=13).value = messageAllInfo[i - 1]['designers']
        sheet.cell(row=i + 1, column=14).value = messageAllInfo[i - 1]['svg_designer']
        sheet.cell(row=i + 1, column=15).value = messageAllInfo[i - 1]['videographers']
        sheet.cell(row=i + 1, column=16).value = messageAllInfo[i - 1]['gif_maker']
        sheet.cell(row=i + 1, column=17).value = messageAllInfo[i - 1]['technical_support']
        sheet.cell(row=i + 1, column=18).value = messageAllInfo[i - 1]['photo_organizer']
        sheet.cell(row=i + 1, column=19).value = messageAllInfo[i - 1]['photo_source']
        sheet.cell(row=i + 1, column=20).value = messageAllInfo[i - 1]['source']
    f.save(u'ZCST公众号文章信息.xls')  # 保存文件


class GUI:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("珠科公众号文章信息抓取器 v23.6.10")

        self.labels = ["platform_UserAgent", "platform_Cookie", "platform_Token", "platform_Fakeid",
                       "WX_User_Agent", "WX_Cookie", "WX_Appmsg_Token", "WX_Pass_ticket", "WX_Key"]
        self.entries = {}

        for idx, label in enumerate(self.labels, start=1):
            tk.Label(self.window, text=label).grid(row=idx, column=0)
            self.entries[label] = tk.Entry(self.window)
            self.entries[label].grid(row=idx, column=1)

        self.btn_submit = tk.Button(self.window, text="提交并运行", command=self.run_main)
        self.btn_submit.grid(row=len(self.labels) + 1, column=0, columnspan=2)

        self.output_area = scrolledtext.ScrolledText(self.window)
        self.output_area.grid(row=len(self.labels) + 2, column=0, columnspan=2)

        self.window.mainloop()

    def run_main(self):
        global platform_UserAgent, platform_Cookie, platform_Token, platform_Fakeid
        global WX_User_Agent, WX_Cookie, WX_Appmsg_Token, WX_Pass_ticket, WX_Key

        platform_UserAgent = self.entries["platform_UserAgent"].get()
        platform_Cookie = self.entries["platform_Cookie"].get()
        platform_Token = self.entries["platform_Token"].get()
        platform_Fakeid = self.entries["platform_Fakeid"].get()
        WX_User_Agent = self.entries["WX_User_Agent"].get()
        WX_Cookie = self.entries["WX_Cookie"].get()
        WX_Appmsg_Token = self.entries["WX_Appmsg_Token"].get()
        WX_Pass_ticket = self.entries["WX_Pass_ticket"].get()
        WX_Key = self.entries["WX_Key"].get()

        # Redirect stdout to the Text widget
        sys.stdout = TextRedirector(self.output_area)

        main()

# Define the TextRedirector class to handle the stdout redirect
class TextRedirector(object):
    def __init__(self, widget):
        self.widget = widget

    def write(self, str):
        self.widget.insert(tk.END, str)
        self.widget.see(tk.END)

    def flush(self):
        pass


if __name__ == '__main__':
    GUI()
