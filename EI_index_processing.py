#-*-coding:utf-8-*-#
from xml.dom.minidom import parse
import xml.dom.minidom
import xlwt
import json
#创建新的工作簿book
book=xlwt.Workbook(encoding='utf-8',style_compression=0)
#添加一个名字为ei-indexing的sheet
sheet=book.add_sheet('ei-indexing',cell_overwrite_ok=True)

# 使用minidom解析器打开XML文档
DOMTree = xml.dom.minidom.parse("ei-engchn.xml")
Data = DOMTree.documentElement
# 在集合中获取所有term
TermInfo = Data.getElementsByTagName("TermInfo")
# 打印每个国家的详细信息

def getText(nodelist):#获取XML中每个nodelist的值
    rc = []
    for node in nodelist:
        if node.nodeType == node.TEXT_NODE:
            rc.append(node.data)
    return ''.join(rc)

def write_excel():#将解析的XML文档写入excel
    sheet.write(row, 0, T)  # T写入第一列
    sheet.write(row, 1, CH)  # CH写入第二列

def addWord(dic,key_word,value):#为字典增加元素
    dic.setdefault(key_word,value)

def Dic_indexing(dic,keys):#输入英文 返回对应的中文
    value=[]
    for k in dic.keys():
        k_split=k.split(',')
        boolen=keys in k_split
        if boolen:
            value.append(dic[k])
    return ''.join(value)

T_CH_dic={}#定义T和CH字典
row=0#定义行号

#写入字典
for Term in TermInfo:
    print "*****TermInfo*****"
    print "写入字典操作processing the :%d rows"%row
    ######第一列##########
    T = Term.getElementsByTagName('T')[0]
    T=T.childNodes[0].data
    ######第二列##########
    ch_list=[]
    for ch in Term.getElementsByTagName('CH'):
        ch_list.append(getText(ch.childNodes))
    CH='#'.join(ch_list)
    addWord(T_CH_dic,T,CH)
    # write_excel()#写入excel
    row+=1

# book.save('ei-engchn.xls')
T_CH_dic=json.dumps(T_CH_dic,ensure_ascii=False)
print T_CH_dic
row=0  #重新定义行号
T_CH_dic=json.loads(T_CH_dic)
for Term in TermInfo:
    print "*****TermInfo*****"
    print "查找回写操作processing the :%d rows"%row
    ######第一列##########
    T1 = Term.getElementsByTagName('T')[0]
    T1 = T1.childNodes[0].data
    ######第二列##########
    ch_list1=[]
    for ch1 in Term.getElementsByTagName('CH'):
        ch_list1.append(getText(ch1.childNodes))
    CH1='#'.join(ch_list1)
    ######第三列##########
    IsPreferred1=Term.getAttribute("type")
    ######第四列##########
    uf_list1=[]
    for uf1 in Term.getElementsByTagName('UF'):
        uf_list1.append(getText(uf1.childNodes))
    UF1='#'.join(uf_list1)
    ######第五列##########
    bt_list1 = []
    for bt1 in Term.getElementsByTagName('BT'):
        bt_list1.append(getText(bt1.childNodes))
    BT1 = '#'.join(bt_list1)
    ######第六列##########
    rt_list1 = []
    for rt1 in Term.getElementsByTagName('RT'):
        rt_list1.append(getText(rt1.childNodes))
    RT1 = '#'.join(rt_list1)
    #######加入NT##############
    nt_list = []
    for nt in Term.getElementsByTagName('NT'):
        nt_list.append(getText(nt.childNodes))
    NT1 = '#'.join(nt_list)
    ######第七列##########
    cl_list1 = []
    for cl1 in Term.getElementsByTagName('CL'):
        cl_list1.append(getText(cl1.childNodes))
    CL1 = '#'.join(cl_list1)
    ######第八列##########
    use_list1 = []
    for use1 in Term.getElementsByTagName('USE'):
        use_list1.append(getText(use1.childNodes))
    USE1 = '#'.join(use_list1)
    ######第九列##########
    useor_list1 = []
    for useor1 in Term.getElementsByTagName('USEOR'):
        useor_list1.append(getText(useor1.childNodes))
    USEOR1 = '#'.join(useor_list1)
    #########################################################
    sheet.write(row, 0, CH1)  # CH1写入第一列
    sheet.write(row,1,T1)  #T写入第二列
    sheet.write(row,2,IsPreferred1) #IsPre写入第三列

    uf11_list=[]
    for uf11 in UF1.split('#'):
        uf11_list.append(Dic_indexing(T_CH_dic,uf11))
    UF1_ch=','.join(uf11_list)
    sheet.write(row, 3, UF1_ch)  # UF_ch写入第四列

    bt11_list = []
    for bt11 in BT1.split('#'):
        bt11_list.append(Dic_indexing(T_CH_dic, bt11))
    BT1_ch = ','.join(bt11_list)
    sheet.write(row, 4, BT1_ch)  # BT_ch写入第五列

    rt11_list = []
    for rt11 in RT1.split('#'):
        rt11_list.append(Dic_indexing(T_CH_dic, rt11))
    RT1_ch = ','.join(rt11_list)
    sheet.write(row, 5, RT1_ch)  # RT_ch写入第六列

    nt11_list = []
    for nt11 in NT1.split('#'):
        nt11_list.append(Dic_indexing(T_CH_dic, nt11))
    NT1_ch = ','.join(nt11_list)
    sheet.write(row, 6, NT1_ch)  # NT_ch写入第7列

    use11_list = []
    for use11 in USE1.split('#'):
        use11_list.append(Dic_indexing(T_CH_dic, use11))
    USE1_ch = ','.join(use11_list)
    sheet.write(row, 7, USE1_ch)  # USE_ch写入第八列

    useor11_list = []
    for useor11 in USEOR1.split('#'):
        useor11_list.append(Dic_indexing(T_CH_dic, useor11))
    USEOR1_ch = ','.join(useor11_list)
    sheet.write(row, 8, USEOR1_ch)  # USEOR_ch写入第八列

    row+=1
book.save('ei-chneng.xls')