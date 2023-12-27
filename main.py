import json
import xml.etree.ElementTree as ET
import openpyxl
import re
import pandas as pd
from gooey import Gooey, GooeyParser

class Element():
    def __init__(self,Elements,dElement):
        self.Elements = Elements
        self.AssembleElement(dElement)


    def AssembleElement(self,edict):
        Element = ET.SubElement(self.Elements, 'Element')

        for key in edict:
            ET.SubElement(Element,key).text = edict[key]

def Classify(cell):
    count = 0
    count_cz = cell.count("L1_LXCZ2_L1_LXCZ2")
    if count_cz > 0:
        count = count_cz + cell.count("LY_") + cell.count("ALM") + cell.count("L3_")+ cell.count("H1_KF")+ cell.count("W1_")
    else:
        count = cell.count("L1_") + cell.count("LY_") + cell.count("ALM") + cell.count("L3_")+ cell.count("H1_KF")+ cell.count("W1_")
    comma = cell.count(",")
    ElementType,ElementName,TagName,Expression,DeviceNum,IsShowTagUnit,DecimalDigits,LoLimit,HiLimit,TagUnit,TextAlign = "","",[],"","","false","2","","","","Center"
    TagName = re.findall("W1_[\w\.]+|L1_[\w\.]+|LY_[\w\.]+|L3_[\w\.]+|ALM-[\w\.]+|H1_KF[\w\.]+", cell)

    if comma > 2:
        IsShowTagUnit = cell.split(',')[2]
        DecimalDigits = "2" if cell.split(',')[3] == "" else str(cell.split(',')[3])
        LoLimit = cell.split(',')[4]
        HiLimit = cell.split(',')[5]
        TagUnit = cell.split(',')[6]
        TextAlign = cell.split(',')[7]

    match count:
        case 0:
            ElementType = "SingleLabel"
            ElementName = "无边框文字"
            if comma == 2:
                ElementType = "Jump"
                ElementName = "跳转"
        case 1:
            ElementType = "DynamicTag"
            ElementName = "动态位号"
            if comma == 1:  # 有逗号，格式为要显示设备名、公式
                DeviceNum = cell.split(',')[0]
                Expression = cell.split(',')[1]
                if DeviceNum != "":
                    ElementType = "DynamicWord"
                    ElementName = "动态文本"
            elif comma == 2:
                ElementType = "Jump"
                ElementName = "跳转"
        case _:
            ElementType = "Statistics"
            ElementName = "统计"
            DeviceNum = cell.split(',')[0]
            Expression = cell.split(',')[1]

            print(cell)
            print(count)

    for i in range(len(TagName)):
        Expression = Expression.replace(TagName[i], "Tag" + str(i + 1))
            # print(f"cell：{cell},Expression:{Expression},DeviceNum:{DeviceNum},IsShowTagUnit:{IsShowTagUnit},DecimalDigits:{DecimalDigits},LoLimit:{LoLimit},HiLimit:{HiLimit}")
    return ElementType,ElementName,TagName,Expression,DeviceNum,IsShowTagUnit,DecimalDigits,LoLimit,HiLimit,TagUnit,TextAlign

# 装配VFTable的cells
def AssembleTable(row,column):
    item = '{"alpha":0,"gbColor":"#FFFFFF"}'
    cells = ''
    for irow in range(row):
        cell = '['
        comma_r = ','
        for icolumn in range(column):
            comma_c = ','
            if icolumn == column - 1:
                comma_c = ''
            cell = cell + item + comma_c
        cell += ']'
        if irow == row - 1:
            comma_r = ''
        cells = cells + cell + comma_r
    cells = "[" + cells + "]"
    return  cells

# 表格和元素参照原点
def Tableinfo(df):
    nrow,ncol = df.shape[0],df.shape[1]
    left,width,ntable = [],0,0
    if nrow <= 27:

        ntable = 1
        match ncol:
            case 5:
                width = 240
            case 6:
                width = 220
            case 7:
                width = 210
            case 8:
                width = 180
            case 9:
                width = 170
            case 10:
                width = 150
            case 11:
                width = 150
            case _:
                width = (1920-30)/ncol
    else:
        ntable = 2
        match ncol:
            case 5:
                width = 162
            case 6:
                width = 140
            case 7:
                width = 129
            case _:
                width = (1920-30)/ncol

    left.append((1920 - width * ncol * ntable) / (ntable+1))
    left.append(2 * left[0] + width * ncol)
    cells = AssembleTable(nrow,ncol)
    print(f"width:{width},ncol:{ncol},表数：{ntable}，left:{left}")
    return (ntable,left, width,ncol,cells)

def ParseXls(df,title):
    tree = ET.parse(r"empty_new.xml")
    root = tree.getroot()
    Elements = root.find("Elements")
    (ntable, left, twidth,ncol,tablecells) = Tableinfo(df)

    mx,my,top = 5,4,80

    # Title
    ElementType,X,Y,Width,Height,Zindex = "SingleLabel",660,16,600,36,1
    edict = {"ElementType":ElementType,"ElementType":ElementType,"X":str(X),"Y":str(Y),"Width":str(Width),"Height":str(Height),"Zindex":str(Zindex),"Alpha":"1","FlowchartID":"3874184148011472","Rotation":"0","TagValueScancycle":"5"}
    edict["OtherAttrs"] = '{"FontFamily": "微软雅黑", "ShowText":"'+ title +'", "FontSize": 36, "TextAlign": "Center","FontColor": 0, "FontWeight": "bold", "orinalId": 3781463406005784}'
    Element(Elements,edict)

    # 标题矩形
    for i in range(ntable):
        Zindex += 1
        ElementType, X, Y, Width, Height, Zindex = "Rect", left[i], top, twidth * ncol, 37, Zindex
        edict = {"ElementType": ElementType, "ElementType": ElementType, "X": str(X), "Y": str(Y), "Width": str(Width),"Height": str(Height), "Zindex": str(Zindex), "Alpha": "0.7", "FlowchartID": "3874184148011472","Rotation": "0", "TagValueScancycle": "5"}
        edict["OtherAttrs"] = '{"WindowWidth":800,"RoateIsUse":false,"RoateSpeed":3000,"GradientColor":"#cccccc","FlowchartNumber":"","BorderColor":"#ffffff","BorderRadius":0,"JumpType":"Jump","Step":10,"BorderWidth":1,"Direction":"0","PageURL":"","IsBack":false,"IsRepeat":false,"FillType":"0","IsUsePreforeAndLast":true,"RoateIsRepeat":false,"IsShowAnimation":false,"Distance":"300","Speed":1000,"FillColor":"#666666","BorderStyle":"0","IsShowBorder":true,"WindowHeight":600,"IsShowFillColor":true,"GradientType":"0","numberStyle":{"colorTag":{"expression":"","showSections":[],"dataType":"true","trueColor":{"color":"#00ff00","id":"0"},"falseColor":{"color":"#ff0000","id":"0"},"tagName":""},"showTag":{"expression":"","showSections":[],"dataType":"true","trueColor":{"switchValue":"true","IsShow":"true","id":"0"},"falseColor":{"switchValue":"false","IsShow":"true","id":"0"},"tagName":""},"trunTag":{"IsBool":"true","minValue":0,"expression":"","IsClockwise":"true","maxValue":100,"maxAngle":180,"dataType":"true","minAngle":0,"IsOn":"false","tagName":"","speed":0},"flashTag":{"expression":"","showSections":[],"dataType":"true","trueColor":{"backColor":"#ff0000","switchValue":"true","id":"0","foreColor":"#00ff00","isFlash":"true","frequency":1},"falseColor":{"backColor":"#ff0000","switchValue":"false","id":"0","foreColor":"#00ff00","isFlash":"false","frequency":1},"tagName":""}},"RoateDirection":0,"orinalId":3880410796378581}'
        Element(Elements, edict)

    # Table
    xpos_list = []
    for i in range(ncol):
        xpos_list.append(i * twidth)
    xpos = str(xpos_list).replace(" ", "")

    for i in range(ntable):
        Zindex += 1
        ElementType,X,Y,Width,Height,Zindex = "FlashTable",left[i],top,twidth*ncol,36,Zindex
        edict = {"ElementType":ElementType,"X":str(X),"Y":str(Y),"Width":str(Width),"Height":"972","Zindex":str(Zindex),"Alpha":"1","FlowchartID":"3874184148011472","Rotation":"0","TagValueScancycle":"5"}
        edict["OtherAttrs"] = '{"border":1,"borderColor":"#ffffff","backgroundColor":"#ffffff","ypos":[0,36,72,108,144,180,216,252,288,324,360,396,432,468,504,540,576,612,648,684,720,756,792,828,864,900,936,972],"cells":[],"xpos":'+xpos+',"backgroundShow":false,"orinalId":3875164117420496}'
        Element(Elements, edict)

    for index,row in df.iterrows():
        itable = 0
        if index > 26:
            itable = 1
        for icol in range(row.size):
            Zindex += 1
            cell = str(row[icol])
            if cell == 'nan':
                continue
            # ElementType,ElementName,TagName,Expression,DeviceNum,IsshowTagUnit,DecimalDigits,LoLimit,HiLimit
            ElementType,ElementName,TagName,Expression,DeviceNum,IsShowTagUnit,DecimalDigits,LoLimit,HiLimit,TagUnit,TextAlign = Classify(cell)
            X = left[itable] + icol*twidth + mx
            Y = top + (index-26*itable)*36 + my
            Width = twidth - 2*mx
            Height = 30 #36 - 2*my
            edict = {"ElementType": ElementType, "X": str(X), "Y": str(Y), "Width": str(Width), "Height": str(Height),"Zindex": str(Zindex), "Alpha": "1", "FlowchartID": "3874184148011472", "Rotation": "0","TagValueScancycle": "5"}

            if ElementType == 'SingleLabel':
                if ',' in cell:
                    ShowText = cell.split(',')[0]
                else:
                    TextAlign = "Left"
                    ShowText = cell
                edict["OtherAttrs"] = '{"FontFamily": "微软雅黑", "ShowText":"' + ShowText + '", "FontSize": 18, "TextAlign": "'+TextAlign+'","FontColor": 0, "FontWeight": false, "orinalId": 3781463406005784,"VerticalAlign":"middle"}'

                # 如果是第一行，且有两个表，多加第二个表的表头
                if index == 0 and ntable == 2:
                    Zindex += 1
                    X = left[1] + icol*twidth + mx
                    edict2 = {"ElementType": ElementType, "X": str(X), "Y": str(Y), "Width": str(Width),"Height": str(Height), "Zindex": str(Zindex), "Alpha": "1", "FlowchartID": "3874184148011472", "Rotation": "0", "TagValueScancycle": "5"}
                    edict2["OtherAttrs"] = '{"FontFamily": "微软雅黑", "ShowText":"' + cell + '", "FontSize": 18, "TextAlign": "Left","FontColor": 0, "FontWeight": false, "orinalId": 3781463406005784,"VerticalAlign":"middle"}'
                    Element(Elements, edict2)
            elif ElementType == 'Jump':
                ShowText = cell.split(',')[0]
                FlowchartNumber = cell.split(',')[2]
                edict["OtherAttrs"] = '{"ShowText": "'+ShowText+'", "WindowWidth": 800, "RoateIsUse": false, "RoateSpeed": 3000,"FlowchartNumber": "'+FlowchartNumber+'", "MutliStatus": [], "JumpType": "Jump", "IsShowBgPic": false,"FontWeight": false, "Direction": "0", "IsShowDefualtTooltipText": true, "PageURL": "","IsBack": false, "IsRepeat": false, "TextAlign": "Left", "RoateIsRepeat": false,"IsShowAnimation": false, "Distance": "300", "BgPicPath": "", "FontFamily": "微软雅黑", "Speed": 1000,"FontColor": "#000000", "WindowHeight": 600, "FontSize": "18", "VTextAlign": "Center","RoateDirection": 0}'
            elif ElementType == "DynamicWord":
                X = X - 2
                Y = Y - 1
                Height = Height - 2
                edict = {"ElementType": ElementType, "X": str(X), "Y": str(Y), "Width": str(Width),
                         "Height": str(Height), "Zindex": str(Zindex), "Alpha": "1",
                         "FlowchartID": "3874184148011472", "Rotation": "0", "TagValueScancycle": "5"}
                edict["OtherAttrs"] = '{"RoateIsUse":false,"RoateSpeed":3000,"MutliStatus":[],"Direction":"0","IsBack":false,"IsRepeat":false,"Status1Text":"'+DeviceNum+'","Status1Style":{"FontFamily":"微软雅黑","FillColor":"#00ff00","FontSize":"18","TextAlign":"center","FontColor":"#000000","BorderColor":"#999999","Alpha":"100%","IsBold":"false","IsShowBorder":"true","BorderWidth":"1","IsShowBG":"true"},"IsUsePreforeAndLast":true,"RoateIsRepeat":false,"IsUseGlobalTagValueScancycle":true,"IsShowTooltip":true,"IsShowAnimation":false,"Distance":"300","Speed":1000,"StatusValue1":"1","StatusValue2":"0","StatusValue3":"-1","Status3Style":{"FontFamily":"微软雅黑","FillColor":"#00ff00","FontSize":"18","TextAlign":"center","FontColor":"#000000","BorderColor":"#505050","Alpha":"100%","IsBold":"false","IsShowBorder":"false","BorderWidth":"1","IsShowBG":"true"},"IsUseGlabolDatasource":true,"Status2Text":"'+DeviceNum+'","Status2Style":{"FontFamily":"微软雅黑","FillColor":"#f0ebf0","FontSize":"18","TextAlign":"center","FontColor":"#000000","BorderColor":"#505050","Alpha":"100%","IsBold":"false","IsShowBorder":"false","BorderWidth":"1","IsShowBG":"false"},"ExceptionStyle":{"FontFamily":"微软雅黑","FillColor":"#f0ebf0","FontSize":"18","TextAlign":"center","FontColor":"#000000","BorderColor":"#505050","Alpha":"100%","IsBold":"false","IsShowBorder":"false","BorderWidth":"1","IsShowBG":"false"},"ExceptionText":"'+DeviceNum+'","TooltipFormat":"","RoateDirection":0,"Status3Text":"'+DeviceNum+'"}'
                edict["TagName"] = TagName[0]
            else:
                if DeviceNum == '': # 动态位号或者统计组件
                    edict["TagName"] = ','.join(TagName)
                    if ElementType == "DynamicTag":
                        edict["OtherAttrs"] = '{"ARExpression":"Tag1&amp;gt;0","FlickFrequency":2,"IsSettledBGWidth":true,"IsUseGlabolStyle":false,"RoateIsUse":false,"IsShowTime":false,"IsshowTagName":false,"ProcessCard":"","IsAlarm":true,"IsshowTagUnit":'+IsShowTagUnit+',"IsFlick":false,"IsAchieveHiHiLimit":false,"isNumberStyle":false,"IsUsePreforeAndLast":true,"IsUseGlobalTagValueScancycle":true,"IsAlarmRestrain":false,"IsshowTagValue":true,"IsShowAnimation":false,"Distance":"300","HiHiStyle":{"FontFamily":"微软雅黑","FillColor":"#ff0000","FontSize":18,"TextAlign":"'+TextAlign+'","Alpha":"100%","BorderColor":"#6600ff","FontColor":"#000000","IsShowBorder":false,"IsBold":false,"BorderWidth":1,"IsShowBG":true},"ValueExceptionShow":"NAN","HiStyle":{"FontFamily":"微软雅黑","FillColor":"#6699cc","FontSize":18,"TextAlign":"'+TextAlign+'","Alpha":"100%","BorderColor":"#6600ff","FontColor":"#ff0000","IsShowBorder":false,"IsBold":false,"BorderWidth":1,"IsShowBG":true},"IsUseGlabolDatasource":true,"IsUseGlobalDecimalDigits":true,"showTagNamePosition":"Up","IsShowExpression":false,"HiHiLimit":"","LoLoStyle":{"FontFamily":"微软雅黑","FillColor":"#ff0000","FontSize":18,"TextAlign":"'+TextAlign+'","Alpha":"100%","BorderColor":"#6600ff","FontColor":"#000000","IsShowBorder":false,"IsBold":false,"BorderWidth":1,"IsShowBG":true},"showTimePosition":"Left","numberStyle":{"colorTag":{"showSections":[],"dataType":"true","trueColor":{"color":"#00ff00","id":"0"},"falseColor":{"color":"#ff0000","id":"0"},"tagName":""},"showTag":{"showSections":[],"dataType":"true","trueColor":{"switchValue":"true","IsShow":"true","id":"0"},"falseColor":{"color":"#ff0000","id":"0"},"tagName":""},"trunTag":{"IsBool":"true","minValue":0,"IsClockwise":"true","maxValue":100,"maxAngle":180,"dataType":"true","minAngle":0,"IsOn":"false","tagName":"","speed":0},"flashTag":{"showSections":[],"dataType":"true","trueColor":{"backColor":"#ff0000","switchValue":"true","id":"0","foreColor":"#00ff00","isFlash":"true","frequency":1},"falseColor":{"color":"#ff0000","id":"0"},"tagName":""}},"IsAchieveLoLoLimit":false,"TooltipFormat":"","RoateDirection":0,"LoLoLimit":"","RoateSpeed":3000,"LoStyle":{"FontFamily":"微软雅黑","FillColor":"#6699cc","FontSize":18,"TextAlign":"'+TextAlign+'","Alpha":"100%","BorderColor":"#6600ff","FontColor":"#ff0000","IsShowBorder":false,"IsBold":false,"BorderWidth":1,"IsShowBG":true},"IsAchieveLoLimit":false,"Direction":"0","AlarmRestrainTagName":"","AlarmConditionType":"Number","ProcessCardSmrId":"","IsBack":false,"IsAchieveHiLimit":false,"IsRepeat":false,"HiLimit":"'+HiLimit+'","alarmId":3991008028838688,"NormalStyle":{"FontFamily":"微软雅黑","FillColor":"#6699cc","FontSize":"18","TextAlign":"center","FontColor":"#000000","BorderColor":"#6600ff","Alpha":"100%","IsBold":"false","IsShowBorder":"false","BorderWidth":"1","IsShowBG":"true"},"RoateIsRepeat":false,"IsShowTooltip":true,"TagNameStyle":"","TagNameHeight":20,"Speed":1000,"IsUseGlobalValueExceptionShow":true,"ColorSections":[],"ProcessCardId":"","TagSource":"TagManage","Text":"","IsShowEffectiveDigit":false,"RegionNumber":3,"IntervalWidth":1,"Expression":"' + Expression + '","DecimalDigits":2,"LoLimit":"'+LoLimit+'","TagUnit":"'+TagUnit+'","orinalId":"e59bf4e88-3f13-7c85-bb6e-003a6bd332f0"}'
                    else:
                        edict["OtherAttrs"] = '{"FlickFrequency": 2, "IsSettledBGWidth": false, "IsUseGlabolStyle": false, "LoLoLimit": "","LoStyle": {"FontFamily": "微软雅黑", "FillColor": "#6699cc", "FontSize": "18","TextAlign": "'+TextAlign+'", "FontColor": "#ff0000", "BorderColor": "#000000","Alpha": "100%", "IsBold": "false", "IsShowBorder": "false", "BorderWidth": "1","IsShowBG": "true"}, "IsAchieveLoLimit": true, "IsAlarm": true, "IsshowTagUnit": '+IsShowTagUnit+',"AlarmConditionType": "Number", "IsFlick": true, "IsAchieveHiLimit": false, "HiLimit": "'+HiLimit+'","IsAchieveHiHiLimit": true,"NormalStyle": "{\\"IsShowBG\\":true,\\"TextAlign\\":\\"'+TextAlign+'\\",\\"FillColor\\":6724044,\\"BorderColor\\":0,\\"BorderWidth\\":1,\\"IsBold\\":false,\\"IsShowBorder\\":false,\\"FontSize\\":18,\\"FontFamily\\":\\"微软雅黑\\",\\"FontColor\\":0,\\"Alpha\\":1}","IsUsePreforeAndLast": true, "IsUseGlobalTagValueScancycle": true, "IsShowTooltip": true,"HiHiStyle": "{\\"IsShowBG\\":true,\\"TextAlign\\":\\"'+TextAlign+'\\",\\"FillColor\\":65535,\\"BorderColor\\":0,\\"BorderWidth\\":1,\\"IsBold\\":false,\\"IsShowBorder\\":false,\\"FontSize\\":18,\\"FontFamily\\":\\"微软雅黑\\",\\"FontColor\\":0,\\"Alpha\\":1}","ValueExceptionShow": "NaN","HiStyle": {"FontFamily": "微软雅黑", "FillColor": "#6699cc", "FontSize": "18","TextAlign": "'+TextAlign+'", "FontColor": "#ff0000", "BorderColor": "#000000","Alpha": "100%", "IsBold": "false", "IsShowBorder": "false", "BorderWidth": "1","IsShowBG": "true"}, "IsUseGlabolDatasource": true, "IsShowEffectiveDigit": false,"IsUseGlobalDecimalDigits": false, "HiHiLimit": "","LoLoStyle": "{\\"IsShowBG\\":true,\\"TextAlign\\":\\"'+TextAlign+'\\",\\"FillColor\\":65535,\\"BorderColor\\":0,\\"BorderWidth\\":1,\\"IsBold\\":false,\\"IsShowBorder\\":false,\\"FontSize\\":18,\\"FontFamily\\":\\"微软雅黑\\",\\"FontColor\\":0,\\"Alpha\\":1}","Expression": "' + Expression + '", "IsFilterExceptionValue": false, "IsAchieveLoLoLimit": true,"LoLimit": "'+LoLimit+'", "DecimalDigits": "'+DecimalDigits+'", "TagUnit": "'+TagUnit+'","TooltipFormat": "名称值时间质量码描述上上限上限下限下下限"}'

            Element(Elements, edict)

    tree.write(title + ".xml", encoding="utf-8")

def ReadExcel(formFile):
    wb = openpyxl.load_workbook(formFile)
    ws = wb.active
    df = pd.read_excel(formFile, ws.title, header=None)
    return (df, ws.title)

def GenerateXML(filename):
    # 读取数采excel文件，获取行数，计表显示的表数
    (df, title) = ReadExcel(filename)
    ParseXls(df, title)

@Gooey(program_description="自动生成数采xml文件。", show_preview_warning=False)
def main():
    # 读取config.json文件，获取上次文件路径等信息
    with open("config/config.json",encoding="utf-8") as f:
        cfg = json.load(f)
    filename = cfg["generate_xml"]["excel"]
    parser = GooeyParser(description="My Cool Gooey App!")
    parser.add_argument('FromFile',help="准备转换的数采Excel文件", widget='FileChooser', default=filename,gooey_options={'wildcard':"Excel file(*.xlsx)|*.xlsx|All files (*.*)|*.*",'message':"pick me"})
    args = parser.parse_args()

    GenerateXML(args.FromFile)

if __name__ == '__main__':
    main()