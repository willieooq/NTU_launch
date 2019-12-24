from linebot.models import *

from openpyxl import load_workbook
#image url
img_map = "https://i.imgur.com/fRf8Rf0.jpg"

v = {'1':'A','2':'B','3':'C','4':'D','5':'E','6':'F'}

def location(area,time):
    if area == 'A':
        if time == '<30':
            return 'A'
        if time == '>30':
            return 'A','B','C','D'
    if area == 'B':
        if time == '<30':
            return 'B'
        if time == '>30':
            return 'A','B','C','E'
    if area == 'C':
        if time == '<30':
            return 'C'
        if time == '>30':
            return 'A','B','C','F'
    if area == 'D':
        if time == '<30':
            return 'A','D'
        if time == '>30':
            return 'A','B','C','D'
    if area == 'E':
        if time == '<30':
            return 'B','E'
        if time == '>30':
            return 'A','B','C','E'
    if area == 'F':
        if time == '<30':
            return 'C','F'
        if time == '>30':
            return 'A','B','C','F'
    if time == 'infinite':
        return 'A','B','C','D','E','F'
#template
#start
in_or_out_template = ButtonsTemplate(
                            title='你現在在哪裡呢?', 
                            text='Please select',
                            actions=[
                            MessageTemplateAction(
                            label="校內", 
                            text="校內"),
                            MessageTemplateAction(
                            label='校外',
                            text='校外')]
						    )
#school area
school = ButtonsTemplate(
                            title='請選擇你在的區域喔~',
                            thumbnail_image_url='https://i.imgur.com/fRf8Rf0.jpg',
                            text='Please select',
                            actions=[
                            MessageTemplateAction(
                            label="A",
                            text="A"),
                            MessageTemplateAction(
                            label="B",
                            text="B"),
                            MessageTemplateAction(
                            label="C",
                            text="C")]
                            )                            
out = ButtonsTemplate(
                            title='請選擇你在的區域喔~',
                            thumbnail_image_url='https://i.imgur.com/fRf8Rf0.jpg',
                            text='Please select',
                            actions=[
                            MessageTemplateAction(
                            label="D",
                            text="D"),
                            MessageTemplateAction(
                            label="E",
                            text="E"),
                            MessageTemplateAction(
                            label="F",
                            text="F")]
                            )
#time
time = ButtonsTemplate(
                            title='你有多少時間呢?',
                            text='Please select',
                            actions=[
                            MessageTemplateAction(
                            label="大於30分鐘",
                            text=">30"),
                            MessageTemplateAction(
                            label="小於30分鐘",
                            text="<30"),
                            MessageTemplateAction(
                            label="我超閒",
                            text="infinite")]
                            )
#price
price_tem = ButtonsTemplate(
                            title='你想吃多好?',
                            text='Please select',
                            actions=[
                            MessageTemplateAction(
                            label="低於100元",
                            text="<100"),
                            MessageTemplateAction(
                            label="100~150元",
                            text="100~150"),
                            MessageTemplateAction(
                            label="老子有錢",
                            text="$$")]
                            )
ok_tem = ButtonsTemplate(
                            title='這家如何?',
                            text='Please select',
                            actions=[
                            MessageTemplateAction(
                            label="就決定是你了",
                            text="Ok"),
                            MessageTemplateAction(
                            label="換一家",
                            text="restart")]
                            )
