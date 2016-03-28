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


def analysis_tj(ls,tab_group=None,lb_tab_top=None):
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
        sales_count =0
        try:
            sales_count = res1.ix[::,7].max()
            print sales_count
        except:
            print 'sales_count is not fetched'
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
        print {'sale':sales_total,'cars':ls1[-2],'people':ls1[-1],'count':sales_count}
        if tab_group is not None:
            res1 = res1.ix[::,:8].copy()
            for i in res1.index:
                if isinstance(res1[res1.columns[5]][i],unicode):
                    v1 = res1[res1.columns[5]][i].strip()
                    if v1==u'特卖':
                        res1.set_value(i,res1.columns[5], u'服装')
                    elif v1==u'生活':
                        res1.set_value(i,res1.columns[5], u'家居生活')
                    elif v1==u'饰品':
                        res1.set_value(i,res1.columns[5], u'配饰')
                    elif v1 not in tab_group.columns:
                        res1.set_value(i,res1.columns[5], u'服装')
                    else:
                        res1.set_value(i,res1.columns[5], v1)
            res2 = res1.ix[:,6].groupby(res1[res1.columns[5]]).sum()#u'group 分析'
            res2 = res2.fillna(0)
            tab_group.set_value(u'天津',col=tab_group.columns,value =tab_group.ix[u'天津']+res2 )
            tab_group.set_value(u'天津',u'销售合计',res2.sum())
        if lb_tab_top is not None:
            res1 = res1.dropna(subset=[res1.columns[3]])
            res1 = res1.sort_values(by =res1.columns[6],ascending=False)
            res1 = res1.set_index(res1.columns[2])
            res1= res1.reindex(columns=[res1.columns[2],res1.columns[3],res1.columns[4],res1.columns[5],res1.columns[6]])
            return res1.ix[:10]
        return {'sale':sales_total,'cars':ls1[-2],'people':ls1[-1],'count':sales_count}
    else:
        print 'no tianjin file in the directory'
def analysis_cy(ls,tab_group=None,lb_tab_top=None):
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
        sales_count =0
        try:
            sales_count = res1.ix[::,7].max()
            print sales_count
        except:
            print 'sales_count is not fetched'
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
        print {'sale':sales_total,'cars':ls1[-2],'people':ls1[-1],'count':sales_count}
        # a4 = axis(name = u'本日总销售',x= res[res.values ==u'本日总销售'].index[0],y= res[res.values ==u'本日总销售'].columns[0])
        # print res.where(res.values ==u'铺位号')
        # print a1.name,a1.x,a1.y,'\n',a2.name,a2.x,a2.y
        if tab_group is not None:
            res1 = res1.ix[::,:8].copy()
            for i in res1.index:
                if isinstance(res1[res1.columns[5]][i],unicode):
                    v1 = res1[res1.columns[5]][i].strip()
                    if v1==u'特卖':
                        res1.set_value(i,res1.columns[5], u'服装')
                    elif v1==u'生活':
                        res1.set_value(i,res1.columns[5], u'家居生活')
                    elif v1==u'饰品':
                        res1.set_value(i,res1.columns[5], u'配饰')
                    elif v1 not in tab_group.columns:
                        res1.set_value(i,res1.columns[5], u'服装')
                    else:
                        res1.set_value(i,res1.columns[5], v1)
            res2 = res1.ix[:,6].groupby(res1[res1.columns[5]]).sum()#u'group 分析'
            res2 = res2.fillna(0)
            tab_group.set_value(u'朝阳',col=tab_group.columns,value =tab_group.ix[u'朝阳']+res2 )
            tab_group.set_value(u'朝阳',u'销售合计',res2.sum())
        if lb_tab_top is not None:
            res1 = res1.dropna(subset=[res1.columns[3]])
            res1 = res1.sort_values(by =res1.columns[6],ascending=False)
            res1 = res1.set_index(res1.columns[2])
            res1= res1.reindex(columns=[res1.columns[2],res1.columns[3],res1.columns[4],res1.columns[5],res1.columns[6]])
            return res1.ix[:10]
        return {'sale':sales_total,'cars':ls1[-2],'people':ls1[-1],'count':sales_count}
    else:
        print 'no chaoyang file in the directory'
def analysis_xd(ls,tab_group=None,lb_tab_top=None):
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
        sales_count =0
        try:
            sales_count = res1.ix[::,7].max()
            print sales_count
        except:
            print 'sales_count is not fetched'
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
        print {'sale':sales_total,'cars':today_cars,'people':today_people,'count':sales_count}
        if tab_group is not None:
            res1 = res1.ix[::,:8].copy()
            for i in res1.index:
                if isinstance(res1[u'品类'][i],unicode):
                    v1 = res1[u'品类'][i].strip()
                    if v1==u'特卖':
                        res1.set_value(i,u'品类', u'服装')
                    elif v1==u'生活':
                        res1.set_value(i,u'品类', u'家居生活')
                    elif v1==u'饰品':
                        res1.set_value(i,u'品类', u'配饰')
                    elif v1 not in tab_group.columns:
                        res1.set_value(i,u'品类', u'服装')
                    else:
                        res1.set_value(i,u'品类', v1)
            res2 = res1.ix[:,6].groupby(res1[u'品类']).sum()#u'group 分析'
            res2 = res2.fillna(0)
            tab_group.set_value(u'西单',col=tab_group.columns,value =tab_group.ix[u'西单']+res2 )
            tab_group.set_value(u'西单',u'销售合计',res2.sum())
        if lb_tab_top is not None:
            res1 = res1.dropna(subset=[res1.columns[3]])
            res1 = res1.sort_values(by =res1.columns[6],ascending=False)
            # res1 = res1.set_index(res1.columns[2])lab_
            return res1.head(10).ix[::,1:7]

            # res1= res1.reindex(columns=[res1.columns[2],res1.columns[3],res1.columns[4],res1.columns[5],res1.columns[6]])
            # lb_tab_top.set_index([u'西单',res1.head(10).ix[::,1:7].columns],inplace = True)
            # lb_tab_top.set_value(u'西单',col=res1.head(10).ix[::,1:7].columns,value =res1.head(10).ix[::,1:7])
            # print lb_tab_top.ix[u'西单']
        return  {'sale':sales_total,'cars':today_cars,'people':today_people,'count':sales_count}
        # a4 = axis(name = u'本日总销售',x= res[res.values ==u'本日总销售'].index[0],y= res[res.values ==u'本日总销售'].columns[0])
        # print res.where(res.values ==u'铺位号')
        # print a1.name,a1.x,a1.y,'\n',a2.name,a2.x,a2.y
    else:
        print 'no xidan file in the directory'
def analysis_sh(ls,tab_group=None,lb_tab_top=None):
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
            if u'品牌名称' in res.columns.tolist():
                print res.columns.tolist().index(u'品牌名称')
            else:
                if res[res.values ==u'品牌'].empty:
                    print res.columns.tolist().index(u'品牌')
                else:
                    a2=axis(name= u'品牌',x=res[res.values ==u'品牌'].index[0],y=res[res.values ==u'品牌'].columns[0])
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
        sales_count = 0
        try:
            sales_count = res1.ix[::,u'交易笔数'].max()
            print sales_count
        except:
            print 'sales_count is not fetched'

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
        print {'sale':sales_total,'cars':ls1[-2],'people':ls1[-1],'count':sales_count}
        if tab_group is not None:
            res1 = res1.ix[::,:9].copy()
            for i in res1.index:
                if isinstance(res1[res1.columns[6]][i],unicode):
                    v1 = res1[res1.columns[6]][i].strip()
                    if v1==u'特卖':
                        res1.set_value(i,res1.columns[6], u'服装')
                    elif v1==u'生活':
                        res1.set_value(i,res1.columns[6], u'家居生活')
                    elif v1==u'饰品':
                        res1.set_value(i,res1.columns[6], u'配饰')
                    elif v1 not in tab_group.columns:
                        res1.set_value(i,res1.columns[6], u'服装')
                    else:
                        res1.set_value(i,res1.columns[6], v1)
            res2 = res1.ix[:,7].groupby(res1[res1.columns[6]]).sum()#u'group 分析'
            res2 = res2.fillna(0)
            tab_group.set_value(u'上海',col=tab_group.columns,value =tab_group.ix[u'上海']+res2 )
            tab_group.set_value(u'上海',u'销售合计',res2.sum())
        if lb_tab_top is not None:
            res1 = res1.dropna(subset=[res1.columns[4]])
            res1 = res1.sort_values(by =res1.columns[7],ascending=False)
            res1 = res1.set_index(res1.columns[2])
            res1= res1.reindex(columns=[res1.columns[3],res1.columns[4],res1.columns[5],res1.columns[6],res1.columns[7]])
            return res1.ix[:10]
        return {'sale':sales_total,'cars':ls1[-2],'people':ls1[-1],'count':sales_count}
        # a4 = axis(name = u'本日总销售',x= res[res.values ==u'本日总销售'].index[0],y= res[res.values ==u'本日总销售'].columns[0])
        # print res.where(res.values ==u'铺位号')
        # print a1.name,a1.x,a1.y,'\n',a2.name,a2.x,a2.y
    else:
        print 'no shanghai file in the directory'
#u'前10行用head',u'p=list.index(value)list为列表的value为查找的值p'
def analysis_sy(ls,tab_group=None,lb_tab_top=None):
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
        sales_count =0
        try:
            sales_count = res1.ix[::,7].max()
            print sales_count
        except:
            print 'sales_count is not fetched'
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
        print {'sale':sales_total,'cars':today_cars,'people':today_people,'count':sales_count}
        if tab_group is not None:
            res1 = res1.ix[::,:8].copy()
            for i in res1.index:
                if isinstance(res1[res1.columns[5]][i],unicode):
                    v1 = res1[res1.columns[5]][i].strip()
                    if v1==u'特卖':
                        res1.set_value(i,res1.columns[5], u'服装')
                    elif v1==u'生活':
                        res1.set_value(i,res1.columns[5], u'家居生活')
                    elif v1==u'饰品':
                        res1.set_value(i,res1.columns[5], u'配饰')
                    elif v1 not in tab_group.columns:
                        res1.set_value(i,res1.columns[5], u'服装')
                    else:
                        res1.set_value(i,res1.columns[5], v1)
            res2 = res1.ix[:,6].groupby(res1[res1.columns[5]]).sum()#u'group 分析'
            res2 = res2.fillna(0)
            tab_group.set_value(u'沈阳',col=tab_group.columns,value =tab_group.ix[u'沈阳']+res2 )
            tab_group.set_value(u'沈阳',u'销售合计',res2.sum())
        if lb_tab_top is not None:
            res1 = res1.dropna(subset=[res1.columns[3]])
            res1 = res1.sort_values(by =res1.columns[6],ascending=False)
            res1 = res1.set_index(res1.columns[2])
            res1= res1.reindex(columns=[res1.columns[2],res1.columns[3],res1.columns[4],res1.columns[5],res1.columns[6]])
            return res1.ix[:10]
        return {'sale':sales_total,'cars':today_cars,'people':today_people,'count':sales_count}
        # a4 = axis(name = u'本日总销售',x= res[res.values ==u'本日总销售'].index[0],y= res[res.values ==u'本日总销售'].columns[0])
        # print res.where(res.values ==u'铺位号')
        # print a1.name,a1.x,a1.y,'\n',a2.name,a2.x,a2.y
    else:
        print 'no shenyang file in the directory'
def analysis_yt(ls,tab_group=None,lb_tab_top=None):
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
        sales_count =0
        try:
            sales_count = res1.ix[::,7].max()
            print sales_count
        except:
            print 'sales_count is not fetched'
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
        print {'sale':sales_total,'cars':ls1[-2],'people':ls1[-1],'count':sales_count}
        if tab_group is not None:
            res1 = res1.ix[::,:8].copy()
            for i in res1.index:
                if isinstance(res1[res1.columns[5]][i],unicode):
                    v1 = res1[res1.columns[5]][i].strip()
                    if v1==u'特卖':
                        res1.set_value(i,res1.columns[5], u'服装')
                    elif v1==u'生活':
                        res1.set_value(i,res1.columns[5], u'家居生活')
                    elif v1==u'饰品':
                        res1.set_value(i,res1.columns[5], u'配饰')
                    elif v1 not in tab_group.columns:
                        res1.set_value(i,res1.columns[5], u'服装')
                    else:
                        res1.set_value(i,res1.columns[5], v1)
            res2 = res1.ix[:,6].groupby(res1[res1.columns[5]]).sum()#u'group 分析'
            res2 = res2.fillna(0)
            tab_group.set_value(u'烟台',col=tab_group.columns,value =tab_group.ix[u'烟台']+res2 )
            tab_group.set_value(u'烟台',u'销售合计',res2.sum())
        if lb_tab_top is not None:
            res1 = res1.dropna(subset=[res1.columns[3]])
            res1 = res1.sort_values(by =res1.columns[6],ascending=False)
            res1 = res1.set_index(res1.columns[2])
            res1= res1.reindex(columns=[res1.columns[2],res1.columns[3],res1.columns[4],res1.columns[5],res1.columns[6]])
            return res1.ix[:10]
        return {'sale':sales_total,'cars':ls1[-2],'people':ls1[-1],'count':sales_count}
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

def create_xls(tab_basic,tab_group,tab_group2,ls_dir_used,tab_top):
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
    tab0.set_value(index = tab_basic.index,col=(u'销售额(万元)',u'当日销售'),value=tab_basic[ls_dir_used[-1]].sale)
    tab0.set_value(index = tab_basic.index,col=(u'销售额(万元)',u'上周同日'),value=tab_basic[ls_dir_used[-7]].sale)
    # tab0.set_value(index = tab_basic.index,col=(u'销售额(万元)',u'增幅'),value=tab_basic[ls_dir_used[-7]].sale)
    tab0.set_value(index = tab_basic.index,col=(u'客流量(万人)',u'当日客流'),value=tab_basic[ls_dir_used[-1]].people)
    tab0.set_value(index = tab_basic.index,col=(u'客流量(万人)',u'上周同日'),value=tab_basic[ls_dir_used[-7]].people)
    tab0.set_value(index = tab_basic.index,col=(u'车流量(车次)',u'当日车流'),value=tab_basic[ls_dir_used[-1]].cars)
    tab0.set_value(index = tab_basic.index,col=(u'车流量(车次)',u'上周同日'),value=tab_basic[ls_dir_used[-7]].cars)
    tab0= tab0.fillna(0)
    tab1 = pa.DataFrame(index = index1,columns=[u'昨日销售',u'今日销售',u'增幅'])
    # print '************','\n',tab_basic.ix[:,ls_dir_used[-1]]['sale']
    tab1.set_value(index=tab1.index,col=u'今日销售',value=tab_basic[ls_dir_used[-1]].sale)
    tab1.set_value(index=tab1.index,col=u'昨日销售',value=tab_basic[ls_dir_used[-2]].sale)
    tab1 = tab1.fillna(0)

    tab2 = pa.DataFrame(index = index1,columns=[u'昨日客流',u'今日客流',u'增幅'])
    tab2.set_value(index=tab2.index,col=u'今日客流',value=tab_basic[ls_dir_used[-1]].people)
    tab2.set_value(index=tab2.index,col=u'昨日客流',value=tab_basic[ls_dir_used[-2]].people)
    tab2 = tab2.fillna(0)

    # u'13日各项目销售额'
    tab3 = pa.DataFrame(data=tab_basic.xs('sale',level='second',axis=1),index = index1,columns=ls_dir_used)
    tab3 = tab3.fillna(0)
    # u'13日各项目客流'
    tab4 = pa.DataFrame(data=tab_basic.xs('people',level='second',axis=1),index = index1,columns=ls_dir_used)
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
    tab6.set_value(index=tab_group.index,col= tab_group.columns,value = tab_group.values)
    # u'上周同日各项目分业态销售状况(万元)'
    tab7 = pa.DataFrame(index = index1,columns=cols)
    tab7 = tab7.fillna(0)
    tab7.set_value(index=tab_group2.index,col= tab_group2.columns,value = tab_group2.values)
    # print tab0,'\r',tab1,'\r',tab2,'\r',tab3,'\r',tab4,'\r',tab5,'\r',tab6,'\r',tab7
    # u'top 10 各项目'
    # tab2 = pa.DataFrame(index = range(1,11),columns=[u'铺位号',u'品牌',u'面积',u'业态',u'当日销售',u'交易笔数',u'客单价',u'上周同日销售',u'销售增幅',u'日坪效',u'备注'])
    # tab2 = tab2.fillna(0)
    # u'13日各项目销售额'
    return [tab0,tab1,tab2,tab3,tab4,tab5,tab6,tab7]
def handle_oneday(tab_basic,dir):
    ls = os.listdir(os.getcwd()+'/'+dir)
    print sys.path[0]+str('/')+str(dir)
    cy=[]#u'朝阳'
    tj=[]#u'天津'
    xd=[]#u'西单'
    cd=[]#u'成都'
    yt=[]#u'烟台'
    xy=[]#u'祥云'
    sh=[]#u'上海'
    sy=[]#u'沈阳
    if  judge_ver()==u'Windows':
        for i in  ls:
            if str(i).decode('gb2312').find(u'朝阳')!= -1:
                cy.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).decode('gb2312').find(u'天津')!= -1 or str(i).decode('gb2312').find(u'大悦城每日销售明细')!= -1 :
                tj.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).decode('gb2312').find(u'西单')!= -1 or str(i).decode('gb2312').find(u'大悦城商户销售')!= -1 and str(i).decode('gb2312').find(u'烟台')==-1:
                xd.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).decode('gb2312').find(u'上海')!= -1:
                sh.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).decode('gb2312').find(u'沈阳')!= -1:
                sy.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).decode('gb2312').find(u'烟台')!= -1:
                yt.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
    else:
        for i in  ls:
            if str(i).find(u'朝阳')!= -1:
                cy.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).find(u'天津')!= -1 or str(i).find(u'大悦城每日销售明细')!= -1:
                tj.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).find(u'西单')!= -1 or str(i).find(u'大悦城商户销售')!= -1 and str(i).find(u'烟台')==-1:
                xd.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).find(u'上海')!= -1:
                sh.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).find(u'沈阳')!= -1:
                sy.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).find(u'烟台')!= -1:
                yt.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
    if len(cy)==0 or len(xd)==0 or len(tj)==0 or len(sh)==0 or len(sy)==0:
        print 'lost file ______ handleoneday()',dir,'\n'
        return
    for i in tab_basic.index.values:
        if i ==u'西单':
            r1=analysis_xd(xd)
            var=[r1['sale'],r1['cars'],r1['people'],r1['count']]
            tab_basic.set_value(i,dir,var)
        elif i ==u'朝阳':
            r1=analysis_cy(cy)
            var=[r1['sale'],r1['cars'],r1['people'],r1['count']]
            tab_basic.set_value(i,dir,var)
        elif i ==u'沈阳':
            r1=analysis_sy(sy)
            var=[r1['sale'],r1['cars'],r1['people'],r1['count']]
            tab_basic.set_value(i,dir,var)
        elif i ==u'上海':
            r1=analysis_sh(sh)
            var=[r1['sale'],r1['cars'],r1['people'],r1['count']]
            tab_basic.set_value(i,dir,var)
        elif i ==u'天津':
            r1=analysis_tj(tj)
            var=[r1['sale'],r1['cars'],r1['people'],r1['count']]
            tab_basic.set_value(i,dir,var)
        elif i ==u'烟台':
            r1=analysis_yt(yt)
            var=[r1['sale'],r1['cars'],r1['people'],r1['count']]
            tab_basic.set_value(i,dir,var)

def handle_oneday_group(tab_group,lb_tab_top,dir):
    ls = os.listdir(os.getcwd()+'/'+dir)
    cy=[]#u'朝阳'
    tj=[]#u'天津'
    xd=[]#u'西单'
    cd=[]#u'成都'
    yt=[]#u'烟台'
    xy=[]#u'祥云'
    sh=[]#u'上海'
    sy=[]#u'沈阳
    if  judge_ver()==u'Windows':
        for i in  ls:
            if str(i).decode('gb2312').find(u'朝阳')!= -1:
                cy.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).decode('gb2312').find(u'天津')!= -1 or str(i).decode('gb2312').find(u'大悦城每日销售明细')!= -1 :
                tj.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).decode('gb2312').find(u'西单')!= -1 or str(i).decode('gb2312').find(u'大悦城商户销售')!= -1 and str(i).decode('gb2312').find(u'烟台')==-1:
                xd.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).decode('gb2312').find(u'上海')!= -1:
                sh.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).decode('gb2312').find(u'沈阳')!= -1:
                sy.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).decode('gb2312').find(u'烟台')!= -1:
                yt.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
    else:
        for i in  ls:
            if str(i).find(u'朝阳')!= -1:
                cy.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).find(u'天津')!= -1 or str(i).find(u'大悦城每日销售明细')!= -1:
                tj.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).find(u'西单')!= -1 or str(i).find(u'大悦城商户销售')!= -1 and str(i).find(u'烟台')==-1:
                xd.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).find(u'上海')!= -1:
                sh.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).find(u'沈阳')!= -1:
                sy.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).find(u'烟台')!= -1:
                yt.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
    if len(cy)==0 or len(xd)==0 or len(tj)==0 or len(sh)==0 or len(sy)==0 or len(yt)==0 :
        print 'lost file ______ handleoneday_group()',dir,'\n'
        return
    for i in tab_group.index.values:
        if i ==u'西单':
            r1=analysis_xd(xd,tab_group,lb_tab_top)
        elif i ==u'朝阳':
            r1=analysis_cy(cy,tab_group,lb_tab_top)
        elif i ==u'沈阳':
            r1=analysis_sy(sy,tab_group,lb_tab_top)
        elif i ==u'上海':
            r1=analysis_sh(sh,tab_group,lb_tab_top)
        elif i ==u'天津':
            r1=analysis_tj(tj,tab_group,lb_tab_top)
        elif i ==u'烟台':
            r1=analysis_yt(yt,tab_group,lb_tab_top)
        #     var=[r1['sale'],r1['cars'],r1['people']]
        #     tab_basic.set_value(i,dir,var)
#u'tab_basic 属于 最终表的部分 结构，主要包含13日销售'
def handle_oneday_top(tab_top,dir):
    # u'暂时被oneday-group 取代'
    ls = os.listdir(os.getcwd()+'/'+dir)
    cy=[]#u'朝阳'
    tj=[]#u'天津'
    xd=[]#u'西单'
    cd=[]#u'成都'
    yt=[]#u'烟台'
    xy=[]#u'祥云'
    sh=[]#u'上海'
    sy=[]#u'沈阳
    if  judge_ver()==u'Windows':
        for i in  ls:
            if str(i).decode('gb2312').find(u'朝阳')!= -1:
                cy.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).decode('gb2312').find(u'天津')!= -1 or str(i).decode('gb2312').find(u'大悦城每日销售明细')!= -1 :
                tj.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).decode('gb2312').find(u'西单')!= -1 or str(i).decode('gb2312').find(u'大悦城商户销售')!= -1 and str(i).decode('gb2312').find(u'烟台')==-1:
                xd.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).decode('gb2312').find(u'上海')!= -1:
                sh.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).decode('gb2312').find(u'沈阳')!= -1:
                sy.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).decode('gb2312').find(u'烟台')!= -1:
                yt.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
    else:
        for i in  ls:
            if str(i).find(u'朝阳')!= -1:
                cy.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).find(u'天津')!= -1 or str(i).find(u'大悦城每日销售明细')!= -1:
                tj.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).find(u'西单')!= -1 or str(i).find(u'大悦城商户销售')!= -1 and str(i).find(u'烟台')==-1:
                xd.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).find(u'上海')!= -1:
                sh.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).find(u'沈阳')!= -1:
                sy.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
            elif str(i).find(u'烟台')!= -1:
                yt.append(sys.path[0]+str('/')+str(dir)+str('/')+i)
    if len(cy)==0 or len(xd)==0 or len(tj)==0 or len(sh)==0 or len(sy)==0 or len(yt)==0 :
        print 'lost file ______ handleoneday_group()',dir,'\n'
        return
    for i in tab_top.columns.levels[0].tolist():
        if i ==u'西单':
            xd_tab_top =analysis_xd(xd,lb_tab_top=tab_top)
        elif i ==u'朝阳':
            print 'ok'
            # r1=analysis_cy(cy,tab_group,lb_tab_top)
        elif i ==u'沈阳':
            print 'ok'
            # r1=analysis_sy(sy,tab_group,lb_tab_top)
        elif i ==u'上海':
            print 'ok'
            # r1=analysis_sh(sh,tab_group,lb_tab_top)
        elif i ==u'天津':
            print 'ok'
            # r1=analysis_tj(tj,tab_group,lb_tab_top)
        elif i ==u'烟台':
            print 'ok'
            # r1=analysis_yt(yt,tab_group,lb_tab_top)

def create_tab_basic():
    print os.getcwd()
    # print os.path.isabs(os.getcwd())
    ls = os.listdir(os.getcwd())
    ls_dir = [] # type:str all dirs included in the current dir
    ls_dir_used= [] # type:str dir used wo choose 13 totally recently
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
    index = [u'西单',u'朝阳',u'沈阳',u'上海',u'天津',u'烟台',u'祥云',u'成都']
    cols1 = ls_dir_used
    cols2= ['sale','cars','people','count']
    multi = pa.MultiIndex.from_product([cols1,cols2], names=['first', 'second'])
    tab_basic = pa.DataFrame(index=index,columns=multi)
    tab_basic = tab_basic.fillna(0)

    cols3= [u'服装',u'配饰',u'化妆品',u'家居生活',u'数码电器',u'皮具',u'正餐',u'非正餐',u'休闲娱乐',u'文教娱乐',u'综合服务',u'专项服务',u'销售合计']
    tab_group = pa.DataFrame(index=index,columns=cols3)
    tab_group = tab_group.fillna(0)
    tab_group2 = tab_group.copy()# u"创建昨日"
    # u'8个项目top10'
    cols_top1 =[u'西单',u'朝阳',u'沈阳',u'上海',u'天津',u'烟台',u'祥云',u'成都']
    cols_top2 = range(1,8)
    # cols_top2= [u'铺位号',u'品牌',u'面积',u'业态',u'当日销售',u'交易笔数',u'客单价',u'上周同日销售',u'销售增幅',u'日坪效',u'备注']
    # multi = pa.MultiIndex.from_product( [cols_top1,range(1,11)], names=['first', 'second'])
    multi = pa.MultiIndex.from_product( [cols_top1,cols_top2], names=['first', 'second'])
    tab_top = pa.DataFrame(index =range(1,11),columns=multi)
    tab_top = tab_top.fillna(0)
    # print tab_basic.xs('sale',level='second',axis=1)
    # for i in ls_dir_used:
    #     handle_oneday(tab_basic,i)
    # handle_oneday_group(tab_group,None,ls_dir_used[-1])# add top and industry
    # handle_oneday_group(tab_group2,None,ls_dir_used[-7])# add top and industry
    handle_oneday_top(tab_top,ls_dir_used[-1])
    handle_oneday_top(tab_top,ls_dir_used[-7])
    return  tab_basic,tab_group,tab_group2,ls_dir_used,tab_top



if __name__ == '__main__':
    judge_ver()
    tab_basic,tab_group,tab_group2,ls_dir_used,tab_top =create_tab_basic()
    # print tab_basic
    ls1= create_xls(tab_basic,tab_group,tab_group2,ls_dir_used,tab_top)
    # data_to_excel(ls1)
    # ''''''test
    # index = [u'西单',u'朝阳',u'沈阳',u'上海',u'天津',u'烟台',u'祥云',u'成都']
    # cols3= [u'服装',u'配饰',u'化妆品',u'家居生活',u'数码电器',u'皮具',u'正餐',u'非正餐',u'休闲娱乐',u'文教娱乐',u'综合服务',u'专项服务',u'销售合计']
    # tab_group = pa.DataFrame(index=index,columns=cols3)
    # tab_group = tab_group.fillna(0)
    # handle_oneday_group(tab_group,True,'1.1')
    # print tab_group