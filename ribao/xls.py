 #coding=utf-8
import pandas as pa
import os
import sys
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
    if ls:
        print ls
        res = pa.read_excel(ls[-1])
        print res.columns
        # print res.index,res.columns
        # ls1 = res[res.values ==u'铺位号'].columns[0]
        a1 = axis()
        a2 = axis()
        a3 = axis()
        if res[res.values ==u'铺位号'].empty:
            print res.columns.tolist().index(u'铺位号')
        else:
            a1 = axis(name= u'铺位号',x=res[res.values ==u'铺位号'].index[0],y=res[res.values ==u'铺位号'].columns[0])
        if res[res.values ==u'品牌名称'].empty:
            print res.columns.tolist().index(u'品牌名称')
        else:
            a2 = axis(name= u'品牌名称',x=res[res.values ==u'品牌名称'].index[0],y=res[res.values ==u'品牌名称'].columns[0])
        if res[res.values ==u'天气'].empty:
            print res.columns.tolist().index(u'天气')
        else:
            a3 = axis(name= u'天气',x=res[res.values ==u'天气'].index[0],y=res[res.values ==u'天气'].columns[0])
        res0 = pa.DataFrame(res.ix[a1.x+1:])
        if (a1.x==a2.x):
            res1 = pa.DataFrame(data=res.ix[a1.x+1:].values,index= res.index[a1.x+1:],columns=res.ix[a1.x,::].tolist())
        else:
            print u'铺位和品牌名称位置不在同一行'
            return
        print res0,res1
        # print res.where(res.values ==u'铺位号')
        # print a1.name,a1.x,a1.y,'\n',a2.name,a2.x,a2.y






    else:
        print 'empty list'

#u'前10行用head',u'p=list.index(value)list为列表的value为查找的值p'
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
    # u'各项目分业态销售状况(万元)'
    cols = [u'服装',u'配饰',u'化妆品',u'家居生活',u'数码电器',u'皮具',u'正餐',u'非正餐',u'休闲娱乐',u'文教娱乐',u'综合服务',u'专项服务',u'销售合计']
    tab5 = pa.DataFrame(index = index1+[u'业态占比',u'上周同日总计',u'环比增幅'],columns=cols)
    tab5 = tab5.fillna(0)
    # u'上周同日各项目分业态销售状况(万元)'
    tab6 = pa.DataFrame(index = index1,columns=cols)
    tab6 = tab6.fillna(0)
    print tab0,'\r',tab1,'\r',tab2,'\r',tab3,'\r',tab4,'\r',tab5,'\r',tab6
    # u'top 10 各项目'
    # tab2 = pa.DataFrame(index = range(1,11),columns=[u'铺位号',u'品牌',u'面积',u'业态',u'当日销售',u'交易笔数',u'客单价',u'上周同日销售',u'销售增幅',u'日坪效',u'备注'])
    # tab2 = tab2.fillna(0)
    # u'13日各项目销售额'
def file_op():
    print os.getcwd()
    # print os.path.isabs(os.getcwd())
    ls = os.listdir(os.getcwd())
    cy=[]#u'朝阳'
    tj=[]#u'天津'
    xd=[]#u'西单'
    cd=[]#u'成都'
    yt=[]#u'烟台'
    xy=[]#u'祥云'
    sh=[]#u'上海'
    sy=[]#u'沈阳'
    for i in  ls:
        if str(i).find(u'朝阳')!= -1:
            cy.append(i)
        elif str(i).find(u'天津')!= -1:
            tj.append(i)
        else:
            print 'empty directory'

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
    analysis_tj(tj)


if __name__ == '__main__':
    judge_ver()
    file_op()
