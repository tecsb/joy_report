 #coding=utf-8
import pandas as pa
import os
import sys
from datetime import datetime
from datetime import timedelta
import re
# reload(sys)
# sys.setdefaultencoding('utf8')
#u'分析天津日报'
def judge_ver():
    import platform
    # print platform.system()
    # print sys.getdefaultencoding()
    if str(platform.system()).encode('utf-8')==u'Windows':
        res = u'Windows'
        #coding=gb2312
    else:
        res = u'Mac'
    reload(sys)
    sys.setdefaultencoding('utf8')
    # print sys.getdefaultencoding()
    return res
class axis:
    def __init__(self, **data):
        self.__dict__.update(data)
            # self.name = ''     # u''
            # self.x = 10     # 尺寸
            # self.y = []     # 列表


def analysis_tj(ls):
    print 'tj report'
    if ls:
        res = pa.read_excel(ls[-1])
        a1 = 0
        a2 = 0
        a3 = 0
        if res[res.values ==u'铺位号'].empty:
            print res.columns.tolist().index(u'铺位号')
        else:
            a1=axis(name= u'铺位号',x=res[res.values ==u'铺位号'].index[0],y=res[res.values ==u'铺位号'].columns[0])
        if res[res.values ==u'品牌名称'].empty:
            print res.columns.tolist().index(u'品牌名称')
        else:
            a2=axis(name= u'品牌名称',x=res[res.values ==u'品牌名称'].index[0],y=res[res.values ==u'品牌名称'].columns[0])
        # if res[res.values ==u'天气'].empty:
        #     print res.columns.tolist().index(u'天气')
        # else:
        #     a3=axis(name= u'天气',x=res[res.values ==u'天气'].index[0],y=res[res.values ==u'天气'].columns[0])
        res0 = pa.DataFrame(res.ix[a1.x+1:])
        if (a1.x==a2.x):
            res1 = pa.DataFrame(data=res.ix[a1.x+1:].values,index= res.index[a1.x+1:],columns=res.ix[a1.x,::].tolist())
        else:
            print u'铺位和品牌名称位置不在同一行'
            return
        sales_total =0
        today_cars=0
        today_people = 0
        for i in res1.ix[::,6].sort_values(ascending=False).dropna():
            if isinstance(i,(float,int)):
                sales_total = i
                break
            else:
                pass
        # if a3==0:
        #     tp = res.columns.tolist()
        #     for i in tp:
        #         if isinstance(i,(int,float)):
        #             pass
        #         elif  isinstance(i,unicode):
        #             if u'车流' in i:
        #                 today_cars=tp[tp.index(i)+1]
        #             elif u'客流'in i:
        #                 today_people = tp[tp.index(i)+1]
        #         else:
        #             pass
        # else:
        #     print 'weather location errors'
        # print {'sale':sales_total,'cars':today_cars,'people':today_people}
        # a4 = axis(name = u'本日总销售',x= res[res.values ==u'本日总销售'].index[0],y= res[res.values ==u'本日总销售'].columns[0])
        # print res.where(res.values ==u'铺位号')
        # print a1.name,a1.x,a1.y,'\n',a2.name,a2.x,a2.y
        if a3==0:
            tp = res.columns.tolist()
            ls1 = []
            for i in tp:
                if isinstance(i,(int,float)):
                    ls1.append(i)
        else:
            print 'weather location errors'
        if(len(ls1)<2):
            print 'error:locate people and cars'
            return
        print {'sale':sales_total,'cars':ls1[-2],'people':ls1[-1]}
    else:
        print 'no tianjin file in the directory'
def analysis_cy(ls):
    print 'cy report'
    if ls:
        res = pa.read_excel(ls[-1])
        a1 = 0
        a2 = 0
        a3 = 0
        if res[res.values ==u'店铺号'].empty:
            print res.columns.tolist().index(u'店铺号')
        else:
            a1=axis(name= u'店铺号',x=res[res.values ==u'店铺号'].index[0],y=res[res.values ==u'店铺号'].columns[0])
        if res[res.values ==u'店铺名称'].empty:
            print res.columns.tolist().index(u'店铺名称')
        else:
            a2=axis(name= u'店铺名称',x=res[res.values ==u'店铺名称'].index[0],y=res[res.values ==u'店铺名称'].columns[0])
        # if res[res.values ==u'天气'].empty:
        #     print res.columns.tolist().index(u'天气')
        # else:
        #     a3=axis(name= u'天气',x=res[res.values ==u'天气'].index[0],y=res[res.values ==u'天气'].columns[0])
        res0 = pa.DataFrame(res.ix[a1.x+1:])
        if (a1.x==a2.x):
            res1 = pa.DataFrame(data=res.ix[a1.x+1:].values,index= res.index[a1.x+1:],columns=res.ix[a1.x,::].tolist())
        else:
            print u'铺位和品牌名称位置不在同一行'
            return
        sales_total =0
        today_cars=0
        today_people = 0
        for i in res1.ix[::,6].sort_values(ascending=False).dropna():
            if isinstance(i,(float,int)):
                sales_total = i;
                break
            else:
                pass
        if a3==0:
            tp = res.columns.tolist()
            ls1 = []
            for i in tp:
                if isinstance(i,(int,float)):
                    ls1.append(i)
        else:
            print 'weather location errors'
        if(len(ls1)<2):
            print 'error:locate people and cars'
            return
        print {'sale':sales_total,'cars':ls1[-2],'people':ls1[-1]}
        # a4 = axis(name = u'本日总销售',x= res[res.values ==u'本日总销售'].index[0],y= res[res.values ==u'本日总销售'].columns[0])
        # print res.where(res.values ==u'铺位号')
        # print a1.name,a1.x,a1.y,'\n',a2.name,a2.x,a2.y
    else:
        print 'no chaoyang file in the directory'
def analysis_xd(ls):
    print 'xd report'
    if ls:
        res = pa.read_excel(ls[-1])
        a1 = 0
        a2 = 0
        a3 = 0 #fault
        if res[res.values ==u'铺位号'].empty:
            print res.columns.tolist().index(u'铺位号')
        else:
            a1=axis(name= u'铺位号',x=res[res.values ==u'铺位号'].index[0],y=res[res.values ==u'铺位号'].columns[0])
        if res[res.values ==u'品牌名称'].empty:
            print res.columns.tolist().index(u'品牌名称')
        else:
            a2=axis(name= u'品牌名称',x=res[res.values ==u'品牌名称'].index[0],y=res[res.values ==u'品牌名称'].columns[0])
        # if res[res.values ==u'天气'].empty:
        #     print res.columns.tolist().index(u'天气')
        # else:
        #     a3=axis(name= u'天气',x=res[res.values ==u'天气'].index[0],y=res[res.values ==u'天气'].columns[0])
        res0 = pa.DataFrame(res.ix[a1.x+1:])
        if (a1.x==a2.x):
            res1 = pa.DataFrame(data=res.ix[a1.x+1:].values,index= res.index[a1.x+1:],columns=res.ix[a1.x,::].tolist())
        else:
            print u'铺位和品牌名称位置不在同一行'
            return
        sales_total =0
        today_cars=0
        today_people = 0
        for i in res1.ix[::,6].sort_values(ascending=False).dropna():
            if isinstance(i,(float,int)):
                sales_total = i
                break
            else:
                pass
        p = re.compile(u'^(\d*)')
        if a3==0:
            tp = res.columns.tolist()
            for i in tp:
                if isinstance(i,(int,float)):
                    pass
                elif  isinstance(i,unicode):
                    if str(i).find(u'车流')>0:
                        tab1 = re.findall('\w*\d+\w*',i)
                        tab2 = re.findall('\w*\d+\.\d+\w*',i)
                        if len(tab2)==0:
                            today_cars=int(tab1[0])
                        else:
                            today_cars=float(tab2[0])
                        # today_cars=tp[tp.index(i)+1]
                    elif str(i).find(u'客流')>0:
                        tab1 = re.findall('\w*\d+\w*',i)
                        tab2 = re.findall('\w*\d+\.\d+\w*',i)
                        if len(tab2)==0:
                            today_people=int(tab1[0])
                        else:
                            today_people=float(tab2[0])
                        # today_people = tp[tp.index(i)+1]
                else:
                    pass
        else:
            print 'weather location errors'
        print {'sale':sales_total,'cars':today_cars,'people':today_people}
        return {'sale':sales_total,'cars':today_cars,'people':today_people}
        # a4 = axis(name = u'本日总销售',x= res[res.values ==u'本日总销售'].index[0],y= res[res.values ==u'本日总销售'].columns[0])
        # print res.where(res.values ==u'铺位号')
        # print a1.name,a1.x,a1.y,'\n',a2.name,a2.x,a2.y
    else:
        print 'no xidan file in the directory'
def analysis_sh(ls):
    print 'sh report'
    if ls:
        res = pa.read_excel(ls[-1])
        a1 = 0
        a2 = 0
        a3 = 0 #fault
        if res[res.values ==u'铺位号'].empty:
            print res.columns.tolist().index(u'铺位号')
        else:
            a1=axis(name= u'铺位号',x=res[res.values ==u'铺位号'].index[0],y=res[res.values ==u'铺位号'].columns[0])
        if res[res.values ==u'品牌名称'].empty:
            print res.columns.tolist().index(u'品牌名称')
        else:
            a2=axis(name= u'品牌名称',x=res[res.values ==u'品牌名称'].index[0],y=res[res.values ==u'品牌名称'].columns[0])
        # if res[res.values ==u'天气'].empty:
        #     print res.columns.tolist().index(u'天气')
        # else:
        #     a3=axis(name= u'天气',x=res[res.values ==u'天气'].index[0],y=res[res.values ==u'天气'].columns[0])
        if (a1.x==a2.x):
            res1 = pa.DataFrame(data=res.ix[a1.x+1:].values,index= res.index[a1.x+1:],columns=res.ix[a1.x,::].tolist())
        else:
            print u'铺位和品牌名称位置不在同一行'
            return
        sales_total =0
        today_cars=0
        today_people = 0
        for i in res1.ix[::,7].sort_values(ascending=False).dropna():
            if isinstance(i,(float,int)):
                sales_total = i
                break
            else:
                pass
        if a3==0:
            tp = res.ix[0,::].tolist()
            ls1 = []
            for i in tp:
                if isinstance(i,(int,float)):
                    ls1.append(i)
        else:
            print 'weather location errors'
        if(len(ls1)<2):
            print 'error:locate people and cars'
            return
        print {'sale':sales_total,'cars':ls1[-2],'people':ls1[-1]}
        # a4 = axis(name = u'本日总销售',x= res[res.values ==u'本日总销售'].index[0],y= res[res.values ==u'本日总销售'].columns[0])
        # print res.where(res.values ==u'铺位号')
        # print a1.name,a1.x,a1.y,'\n',a2.name,a2.x,a2.y
    else:
        print 'no shanghai file in the directory'
#u'前10行用head',u'p=list.index(value)list为列表的value为查找的值p'
def analysis_sy(ls):
    print 'sy report'
    if ls:
        res = pa.read_excel(ls[-1])
        a1 = 0
        a2 = 0
        a3 = 0 #fault
        if res[res.values ==u'铺位号'].empty:
            print res.columns.tolist().index(u'铺位号')
        else:
            a1=axis(name= u'铺位号',x=res[res.values ==u'铺位号'].index[0],y=res[res.values ==u'铺位号'].columns[0])
        if res[res.values ==u'品牌名称'].empty:
            print res.columns.tolist().index(u'品牌名称')
        else:
            a2=axis(name= u'品牌名称',x=res[res.values ==u'品牌名称'].index[0],y=res[res.values ==u'品牌名称'].columns[0])
        # if res[res.values ==u'天气'].empty:
        #     print res.columns.tolist().index(u'天气')
        # else:
        #     a3=axis(name= u'天气',x=res[res.values ==u'天气'].index[0],y=res[res.values ==u'天气'].columns[0])
        if (a1.x==a2.x):
            res1 = pa.DataFrame(data=res.ix[a1.x+1:].values,index= res.index[a1.x+1:],columns=res.ix[a1.x,::].tolist())
        else:
            print u'铺位和品牌名称位置不在同一行'
            return
        sales_total =0
        today_cars=0
        today_people = 0
        for i in res1.ix[::,6].sort_values(ascending=False).dropna():
            if isinstance(i,(float,int)):
                sales_total = i
                break
            else:
                pass
        if a3==0:
            tp = res.columns.tolist()
            for i in tp:
                if isinstance(i,(int,float)):
                    pass
                elif  isinstance(i,unicode):
                    if str(i).find(u'车流')>0:
                        next = tp[tp.index(i)+1]
                        if isinstance(next,(int,float)):
                            today_cars = next
                        else:
                            tab1 = re.findall('\w*\d+\w*',next)
                            if len(tab1)==1:
                                today_cars = tab1[0]
                            elif len(tab1)>1:
                                today_cars = float(tab1[-1])+(float(tab1[-2])*1000)
                    elif str(i).find(u'客流')>0:
                        next =tp[tp.index(i)+1]
                        if isinstance(next,(int,float)):
                            today_people = next
                        else:
                            tab1 = re.findall('\w*\d+\w*',i)
                            if len(tab1)==1:
                                today_cars = tab1[0]
                            elif len(tab1)>1:
                                today_cars = float(tab1[-1])+(float(tab1[-2])*1000)
                else:
                    pass
        else:
            print 'weather location errors'
        print {'sale':sales_total,'cars':today_cars,'people':today_people}
        # a4 = axis(name = u'本日总销售',x= res[res.values ==u'本日总销售'].index[0],y= res[res.values ==u'本日总销售'].columns[0])
        # print res.where(res.values ==u'铺位号')
        # print a1.name,a1.x,a1.y,'\n',a2.name,a2.x,a2.y
    else:
        print 'no shenyang file in the directory'
def analysis_yt(ls):
    print 'yt report'
    if ls:
        res = pa.read_excel(ls[-1])
        a1 = 0
        a2 = 0
        a3 = 0
        if res[res.values ==u'铺位号'].empty:
            print res.columns.tolist().index(u'铺位号')
        else:
            a1=axis(name= u'铺位号',x=res[res.values ==u'铺位号'].index[0],y=res[res.values ==u'铺位号'].columns[0])
        if res[res.values ==u'品牌名称'].empty:
            print res.columns.tolist().index(u'品牌名称')
        else:
            a2=axis(name= u'品牌名称',x=res[res.values ==u'品牌名称'].index[0],y=res[res.values ==u'品牌名称'].columns[0])
        # if res[res.values ==u'天气'].empty:
        #     print res.columns.tolist().index(u'天气')
        # else:
        #     a3=axis(name= u'天气',x=res[res.values ==u'天气'].index[0],y=res[res.values ==u'天气'].columns[0])
        res0 = pa.DataFrame(res.ix[a1.x+1:])
        if (a1.x==a2.x):
            res1 = pa.DataFrame(data=res.ix[a1.x+1:].values,index= res.index[a1.x+1:],columns=res.ix[a1.x,::].tolist())
        else:
            print u'铺位和品牌名称位置不在同一行'
            return
        sales_total =0
        today_cars=0
        today_people = 0
        for i in res1.ix[::,6].sort_values(ascending=False).dropna():
            if isinstance(i,(float,int)):
                sales_total = i;
                break
            else:
                pass
        if a3==0:
            tp = res.columns.tolist()
            ls1 = []
            for i in tp:
                if isinstance(i,(int,float)):
                    ls1.append(i)
        else:
            print 'weather location errors'
        if(len(ls1)<2):
            print 'error:locate people and cars'
            return
        print {'sale':sales_total,'cars':ls1[-2],'people':ls1[-1]}
        # a4 = axis(name = u'本日总销售',x= res[res.values ==u'本日总销售'].index[0],y= res[res.values ==u'本日总销售'].columns[0])
        # print res.where(res.values ==u'铺位号')
        # print a1.name,a1.x,a1.y,'\n',a2.name,a2.x,a2.y
    else:
        print 'no chaoyang file in the directory'
def data_to_excel(lists):
    writer = pa.ExcelWriter('test.xlsx',engine='xlsxwriter')   # Creating Excel Writer Object from Pandas
    # workbook=writer.book
    row = 0
    for i in lists:
        i.to_excel(writer,sheet_name='Validation',startrow=row, startcol=0)
        row = row +i.shape[0]+len(i.columns.names) + 5
    writer.save()

def create_xls():
    # res2 = pa.read_excel('cy.xlsx')

    title1 = pa.DataFrame(data=[u'朝阳前十'],index=range(1))
    title2 = pa.Series(data=[u'朝阳前十'],index=range(14))
    '''
    ls1 = [u'销售额(万元)',u'客流量(万人)',u'车流量(车次)']
    ls2 = [u'当日销售',u'上周同日',u'增幅',u'销售占比']
    multi = pa.MultiIndex.from_product([ls1,ls2], names=['first', 'second'])
    print multi
    '''
    ls1 = [u'销售额(万元)',u'销售额(万元)',u'销售额(万元)',u'销售额(万元)',u'客流量(万人)',u'客流量(万人)',u'客流量(万人)',u'客流量(万人)',u'车流量(车次)',u'车流量(车次)',u'车流量(车次)',u'车流量(车次)']
    ls2 = [u'当日销售',u'上周同日',u'增幅',u'销售占比',u'当日客流',u'上周同日',u'增幅',u'客流占比',u'当日车流',u'上周同日',u'增幅',u'车流占比']
    multi = [ls1,ls2]
    index1 = [u'西单',u'朝阳',u'沈阳',u'上海',u'天津',u'烟台',u'祥云',u'成都',u'各项目合计']
    tab0 = pa.DataFrame(index= index1,columns= multi)
    print tab0.shape
    tab0= tab0.fillna(0)

    tab1 = pa.DataFrame(index = index1,columns=[u'昨日销售',u'今日销售',u'增幅'])
    tab1 = tab1.fillna(0)

    tab2 = pa.DataFrame(index = index1,columns=[u'昨日客流',u'今日客流',u'增幅'])
    tab2 = tab1.fillna(0)

    # u'13日各项目销售额'
    tab3 = pa.DataFrame(index = index1,columns=range(13))
    tab3 = tab3.fillna(0)
    # u'13日各项目客流'
    tab4 = pa.DataFrame(index = index1,columns=range(13))
    tab4 = tab4.fillna(0)
    #u'项目客群指标列表'
    ls3 = [u'提袋率',u'提袋率',u'提袋率',u'客单价',u'客单价',u'客单价',u'车/客流占比',u'车/客流占比',u'车/客流占比',u'车位转换率',u'车位转换率',u'车位转换率']
    ls4 = [u'当日',u'上周同日',u'增幅',u'当日',u'上周同日',u'增幅',u'车流量',u'客流量',u'车/客流占比',u'车流量',u'车位数',u'车位转换率']
    multi5 = [ls3,ls4]
    index5 = [u'西单',u'朝阳',u'沈阳',u'上海',u'天津',u'烟台',u'祥云',u'成都',u'平均']
    tab5 = pa.DataFrame(index= index5,columns= multi5)
    tab5 = tab5.fillna(0)
    # u'各项目分业态销售状况(万元)'
    cols = [u'服装',u'配饰',u'化妆品',u'家居生活',u'数码电器',u'皮具',u'正餐',u'非正餐',u'休闲娱乐',u'文教娱乐',u'综合服务',u'专项服务',u'销售合计']
    tab6 = pa.DataFrame(index = index1+[u'业态占比',u'上周同日总计',u'环比增幅'],columns=cols)
    tab6 = tab6.fillna(0)
    # u'上周同日各项目分业态销售状况(万元)'
    tab7 = pa.DataFrame(index = index1,columns=cols)
    tab7 = tab7.fillna(0)
    print tab0,'\r',tab1,'\r',tab2,'\r',tab3,'\r',tab4,'\r',tab5,'\r',tab6,'\r',tab7
    # u'top 10 各项目'
    # tab2 = pa.DataFrame(index = range(1,11),columns=[u'铺位号',u'品牌',u'面积',u'业态',u'当日销售',u'交易笔数',u'客单价',u'上周同日销售',u'销售增幅',u'日坪效',u'备注'])
    # tab2 = tab2.fillna(0)
    # u'13日各项目销售额'
    data_to_excel([tab0,tab1,tab2,tab3,tab4,tab5,tab6,tab7])
def file_op():
    print os.getcwd()
    # print os.path.isabs(os.getcwd())
    ls = os.listdir(os.getcwd())
    ls_dir = [] # type:str all dirs included in the current dir
    ls_dir_used= [] # type:str dir used wo choose 13 totaly
    ls_date = [] # type:datetime
    for i in ls:
        if os.path.isdir(i):
            ls_dir.append(i)
            if len(re.findall('\w*\d+.\d+\w*',i))>0:
                tmp = datetime.strptime(i,'%m.%d')
                if tmp.month==1:
                    ls_date.append(datetime(tmp.year+1,tmp.month,tmp.day))
                else:
                    ls_date.append(tmp)
            else:
                print u'文件夹的日期格式有误'
    ls_date = sorted(ls_date)
    try:
        str_date_index = [i.strftime('%m.%d') for i in ls_date[-13:]]
        for i in str_date_index:
            tmp =re.findall('[1-9]+[0-9]*\.[1-9]+[0-9]*',i)
            if len(tmp)>0:
                ls_dir_used.append(tmp[0])
            else:
                pat = re.compile('0')
                ls_dir_used.append(pat.sub('',i))
    except:
        print u'取值超出索引序列，文件夹数量不足'
    cy=[]#u'朝阳'
    tj=[]#u'天津'
    xd=[]#u'西单'
    cd=[]#u'成都'
    yt=[]#u'烟台'
    xy=[]#u'祥云'
    sh=[]#u'上海'
    sy=[]#u'沈阳'
    if  judge_ver()==u'Windows':
        for i in  ls:
            if str(i).decode('gb2312').find(u'朝阳')!= -1:
                cy.append(i)
            elif str(i).decode('gb2312').find(u'天津')!= -1:
                tj.append(i)
            elif str(i).decode('gb2312').find(u'西单')!= -1:
                xd.append(i)
            elif str(i).decode('gb2312').find(u'上海')!= -1:
                sh.append(i)
            elif str(i).decode('gb2312').find(u'沈阳')!= -1:
                sy.append(i)
            elif str(i).decode('gb2312').find(u'烟台')!= -1:
                yt.append(i)
    else:
        for i in  ls:
            if str(i).find(u'朝阳')!= -1:
                cy.append(i)
            elif str(i).find(u'天津')!= -1:
                tj.append(i)
            elif str(i).find(u'西单')!= -1:
                xd.append(i)
            elif str(i).find(u'上海')!= -1:
                sh.append(i)
            elif str(i).find(u'沈阳')!= -1:
                sy.append(i)
            elif str(i).find(u'烟台')!= -1:
                yt.append(i)
    index = [u'西单',u'朝阳',u'沈阳',u'上海',u'天津',u'烟台',u'祥云',u'成都']
    cols1 = ls_dir_used
    cols2= ['sale','cars','people']

    multi = pa.MultiIndex.from_product([cols1,cols2], names=['first', 'second'])
    tab_basic = pa.DataFrame(index=index,columns=multi)
    tab_basic = tab_basic.fillna(0)
    tab_basic.set_value(u'西单','1.1',analysis_xd(ls).values())
    '''= analysis_xd(xd)['cars']'''
        # if isinstance(i,unicode) and judge_ver()==u'Mac':
        #     if u'朝阳' in i:
        #         cy.append(i)
        #     if u'天津' in i:
        #         tj.append(i)
        # elif isinstance(i, str) and judge_ver()==u'Windows':
        #     if u'朝阳' in i.decode('gbk').encode('utf-8'):
        #         cy.append(i)
        #     if u'天津' in i.decode('gbk').encode('utf-8'):
        #         tj.append(i)
    # analysis_tj(tj)
    # analysis_cy(cy)
    # analysis_xd(xd)
    # analysis_sh(sh)
    # analysis_sy(sy)
    # analysis_yt(yt)
def create_basic_tab(t1,t2):
    tab_basic = pa.DataFrame(index=t1,columns=t2)
    tab_basic = tab_basic.fillna(0)
    print tab_basic



if __name__ == '__main__':
    judge_ver()
    file_op()
    # create_xls()
