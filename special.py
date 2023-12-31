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
    count = cell.count("L1_") + cell.count("LY_") + cell.count("ALM") + cell.count("L3_")+ cell.count("H1_")+ cell.count("W1_")
    comma = cell.count(",")
    ElementType,ElementName,TagName,Expression,DeviceNum,IsShowTagUnit,DecimalDigits,LoLimit,HiLimit,TagUnit,TextAlign = "","",[],"","","false","2","","","","Center"
    match count:
        case 0:
            ElementType = "SingleLabel"
            ElementName = "无边框文字"
        case _:
            ElementType = "Statistics"
            ElementName = "统计"
            TagName = re.findall("W1_[\w\.]+|L1_[\w\.]+|LY_[\w\.]+|L3_[\w\.]+|ALM-[\w\.]+|H1_[\w\.]+", cell)
            DeviceNum = cell.split(',')[0]
            Expression = cell.split(',')[1]
            if comma == 1:  # 有逗号，格式为要显示设备名、公式（单Tag也作为公式）、是否显示单位、小数点位数、低限报警、高限报警，文本对齐
                if DeviceNum != "":
                    ElementType = "DynamicWord"
                    ElementName = "动态文本"

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
    return (ntable,left, width,ncol,cells)

def ParseXls(df,title):
    tree = ET.parse(r"empty_new.xml")
    root = tree.getroot()
    Elements = root.find("Elements")

    mx,my,left,top,twidth = 2,1,70,30,100

    # Title
    ElementType,X,Y,Width,Height,Zindex = "SingleLabel",660,16,600,36,1
    edict = {"ElementType":ElementType,"ElementType":ElementType,"X":str(X),"Y":str(Y),"Width":str(Width),"Height":str(Height),"Zindex":str(Zindex),"Alpha":"1","FlowchartID":"3874184148011472","Rotation":"0","TagValueScancycle":"5"}
    edict["OtherAttrs"] = '{"FontFamily": "微软雅黑", "ShowText":"'+ title +'", "FontSize": 36, "TextAlign": "Center","FontColor": 0, "FontWeight": "bold", "orinalId": 3781463406005784}'
    Element(Elements,edict)

    for index,row in df.iterrows():
        for icol in range(row.size):
            Zindex += 1
            cell = str(row[icol])
            if cell == 'nan':
                continue
            # ElementType,ElementName,TagName,Expression,DeviceNum,IsshowTagUnit,DecimalDigits,LoLimit,HiLimit
            ElementType,ElementName,TagName,Expression,DeviceNum,IsShowTagUnit,DecimalDigits,LoLimit,HiLimit,TagUnit,TextAlign = Classify(cell)
            X = left + icol*twidth + mx
            Y = top + index*28 + my
            Width = twidth - 2*mx
            Height = 28 - 2*my
            edict = {"ElementType": ElementType, "X": str(X), "Y": str(Y), "Width": str(Width), "Height": str(Height),"Zindex": str(Zindex), "Alpha": "1", "FlowchartID": "3874184148011472", "Rotation": "0","TagValueScancycle": "5"}
            if ElementType == 'SingleLabel':
                Zindex += 1
                edict["OtherAttrs"] = '{"FontFamily": "微软雅黑", "ShowText":"' + cell + '", "FontSize": 17, "TextAlign": "Left","FontColor": 0, "FontWeight": false, "orinalId": 3781463406005784,"VerticalAlign":"middle"}'
                Element(Elements, edict)
            else:
                print(f"TextAlign:{TextAlign}")
                if DeviceNum == '':
                    Zindex += 1
                    edict["OtherAttrs"] = '{"FlickFrequency": 2, "IsSettledBGWidth": false, "IsUseGlabolStyle": false, "LoLoLimit": "","LoStyle": {"FontFamily": "微软雅黑", "FillColor": "#6699cc", "FontSize": "17","TextAlign": "'+TextAlign+'", "FontColor": "#ff0000", "BorderColor": "#000000","Alpha": "100%", "IsBold": "false", "IsShowBorder": "false", "BorderWidth": "1","IsShowBG": "true"}, "IsAchieveLoLimit": true, "IsAlarm": true, "IsshowTagUnit": '+IsShowTagUnit+',"AlarmConditionType": "Number", "IsFlick": true, "IsAchieveHiLimit": false, "HiLimit": "'+HiLimit+'","IsAchieveHiHiLimit": true,"NormalStyle": "{\\"IsShowBG\\":true,\\"TextAlign\\":\\"'+TextAlign+'\\",\\"FillColor\\":6724044,\\"BorderColor\\":0,\\"BorderWidth\\":1,\\"IsBold\\":false,\\"IsShowBorder\\":false,\\"FontSize\\":18,\\"FontFamily\\":\\"微软雅黑\\",\\"FontColor\\":0,\\"Alpha\\":1}","IsUsePreforeAndLast": true, "IsUseGlobalTagValueScancycle": true, "IsShowTooltip": true,"HiHiStyle": "{\\"IsShowBG\\":true,\\"TextAlign\\":\\"'+TextAlign+'\\",\\"FillColor\\":65535,\\"BorderColor\\":0,\\"BorderWidth\\":1,\\"IsBold\\":false,\\"IsShowBorder\\":false,\\"FontSize\\":18,\\"FontFamily\\":\\"微软雅黑\\",\\"FontColor\\":0,\\"Alpha\\":1}","ValueExceptionShow": "NaN","HiStyle": {"FontFamily": "微软雅黑", "FillColor": "#6699cc", "FontSize": "17","TextAlign": "'+TextAlign+'", "FontColor": "#ff0000", "BorderColor": "#000000","Alpha": "100%", "IsBold": "false", "IsShowBorder": "false", "BorderWidth": "1","IsShowBG": "true"}, "IsUseGlabolDatasource": true, "IsShowEffectiveDigit": false,"IsUseGlobalDecimalDigits": false, "HiHiLimit": "","LoLoStyle": "{\\"IsShowBG\\":true,\\"TextAlign\\":\\"'+TextAlign+'\\",\\"FillColor\\":65535,\\"BorderColor\\":0,\\"BorderWidth\\":1,\\"IsBold\\":false,\\"IsShowBorder\\":false,\\"FontSize\\":17,\\"FontFamily\\":\\"微软雅黑\\",\\"FontColor\\":0,\\"Alpha\\":1}","Expression": "' + Expression + '", "IsFilterExceptionValue": false, "IsAchieveLoLoLimit": true,"LoLimit": "'+LoLimit+'", "DecimalDigits": "'+DecimalDigits+'", "TagUnit": "'+TagUnit+'","TooltipFormat": "名称值时间质量码描述上上限上限下限下下限"}'
                    edict["TagName"] = ','.join(TagName)
                    Element(Elements, edict)
                else:
                    Zindex += 1
                    X = X - 2
                    Y = Y - 1
                    Height = Height - 2
                    edict = {"ElementType": ElementType, "X": str(X), "Y": str(Y), "Width": str(Width),
                             "Height": str(Height), "Zindex": str(Zindex), "Alpha": "1",
                             "FlowchartID": "3874184148011472", "Rotation": "0", "TagValueScancycle": "5"}
                    edict["OtherAttrs"] = '{"RoateIsUse":false,"RoateSpeed":3000,"MutliStatus":[],"Direction":"0","IsBack":false,"IsRepeat":false,"Status1Text":"'+DeviceNum+'","Status1Style":{"FontFamily":"微软雅黑","FillColor":"#00ff00","FontSize":"17","TextAlign":"center","FontColor":"#000000","BorderColor":"#999999","Alpha":"100%","IsBold":"false","IsShowBorder":"true","BorderWidth":"1","IsShowBG":"true"},"IsUsePreforeAndLast":true,"RoateIsRepeat":false,"IsUseGlobalTagValueScancycle":true,"IsShowTooltip":true,"IsShowAnimation":false,"Distance":"300","Speed":1000,"StatusValue1":"1","StatusValue2":"0","StatusValue3":"-1","Status3Style":{"FontFamily":"微软雅黑","FillColor":"#00ff00","FontSize":"17","TextAlign":"center","FontColor":"#000000","BorderColor":"#505050","Alpha":"100%","IsBold":"false","IsShowBorder":"false","BorderWidth":"1","IsShowBG":"true"},"IsUseGlabolDatasource":true,"Status2Text":"'+DeviceNum+'","Status2Style":{"FontFamily":"微软雅黑","FillColor":"#f0ebf0","FontSize":"17","TextAlign":"center","FontColor":"#000000","BorderColor":"#505050","Alpha":"100%","IsBold":"false","IsShowBorder":"false","BorderWidth":"1","IsShowBG":"false"},"ExceptionStyle":{"FontFamily":"微软雅黑","FillColor":"#f0ebf0","FontSize":"17","TextAlign":"center","FontColor":"#000000","BorderColor":"#505050","Alpha":"100%","IsBold":"false","IsShowBorder":"false","BorderWidth":"1","IsShowBG":"false"},"ExceptionText":"'+DeviceNum+'","TooltipFormat":"","RoateDirection":0,"Status3Text":"'+DeviceNum+'"}'
                    edict["TagName"] = TagName[0]
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