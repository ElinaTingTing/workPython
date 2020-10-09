import os
import pandas as pd
import datetime as dt
import numpy as np
import re
from dateutil.relativedelta import relativedelta  
from dateutil.rrule import *


path=os.getcwd()

time='2020_10_09'
data=pd.read_excel(path+r'\data\creditTotalAmountResult.xls')


def amountformat (amount):
    if amount<=2000:
        return '0-2000'
    elif amount <=2500:
        return '2000-2500'
    elif amount <= 3500:
        return '2500-3500'
    elif amount <=5000:
        return '3500-5000'
    elif amount <=7500:
        return '5000-7500'
    elif amount <=9500:
        return '7500-9500'
    elif amount <= 12000:
        return '9500-12000'
    elif amount <= 18000:
        return '12000-18000'
    else:
        return '18000+'  
def rateformat (rate):
    if rate==0:
        return '0'
    elif rate<=0.1:
        return '0-0.1'
    elif rate <=0.2:
        return '0.1-0.2'
    elif rate <= 0.3:
        return '0.2-0.3'
    elif rate <=0.4:
        return '0.3-0.4'
    elif rate <=0.5:
        return '0.4-0.5'
    elif rate <= 0.6:
        return '0.5-0.6'
    elif rate <= 0.7:
        return '0.6-0.7'
    elif rate <= 0.8:
        return '0.7-0.8'
    elif rate <= 0.9:
        return '0.8-0.9'
    else:
        return '0.9-1'
def duedays (day):
    if day==0:
        return '0'

    elif day <= 31:
        return '0-30'
    elif day <=90:
        return '30-90'
    else:
        return '90+'
		
		
		

data1=data.copy()

data1['授信时间']=data1['授信时间'].apply(lambda x: dt.datetime.strptime(x ,'%Y-%m-%d %H:%M:%S' ))
data1['授信金额_f']=data1['授信金额'].apply(lambda x: amountformat(x))
data1['逾期天数_f']=data1['逾期天数'].apply(lambda x: duedays(x))
data1['使用率']=data1['已用额度']/data1['授信金额']
data1['使用率_f']=data1['使用率'].apply(lambda x: rateformat(x))
data1['类型']=data1['类型'].apply(lambda x: '增量' if x=='一键激活' else x)


dd=data1.groupby(['授信金额_f','类型'])['寺库用户ID'].count().reset_index().rename({'寺库用户ID':'数量'},axis=1)
tt=dd.pivot(columns='类型' , values='数量' , index='授信金额_f')
mm=tt.reset_index()
mm['sort']= mm['授信金额_f'].apply(lambda x: int(re.findall('^\d+',x)[0]))
output=mm.sort_values(by='sort',ascending=True)
output=output.rename({'授信金额_f':'授信金额'},axis=1)
output_amount_hist=output.reset_index().fillna(0).drop(columns=['sort','index'])


output_amount_hist['总计']='总计'
col=list(output_amount_hist.columns.values)
output_amount_hist.columns=col
col1=list(output_amount_hist.columns.values)
col1.remove('总计')
tt=output_amount_hist.groupby('总计')[col1].sum().reset_index()
tt=tt.rename({'总计':'授信金额'},axis=1)
output_amount_hist=output_amount_hist.append(tt).drop(columns='总计')[col1]


output_amount=output_amount_hist.copy()
col2=list(output_amount.columns.values)
dic2=dict(output_amount[col2].max())

for value in col2:
    try:
        dic2[value]>0
        num=dic2[value]
        output_amount[value]=output_amount[value]/num
        output_amount[value]=output_amount.apply(lambda x: "%.2f%%" % (x[value] * 100) ,axis=1)
    except:
        print(value)
        
        
pct=output_amount.set_index(['授信金额']).T
amt=output_amount_hist.set_index(['授信金额']).T

pct['统计类型']='百分比'
amt['统计类型']='账户数'

output_amount_total=pd.concat([pct,amt])
output_amount_hist=output_amount_total.reset_index().set_index(['index','统计类型']).sort_index().T
output_amount_hist.columns.names=['业务类型','统计类型']

dd=data1.groupby(['使用率_f','类型'])['寺库用户ID'].count().reset_index().rename({'寺库用户ID':'数量'},axis=1)
tt=dd.pivot(columns='类型' , values='数量' , index='使用率_f')
mm=tt.reset_index()
mm['sort']= mm['使用率_f'].apply(lambda x: int(re.findall('^\d+',x)[0]))
output=mm.sort_values(by='sort',ascending=True)
output=output.rename({'使用率_f':'使用率'},axis=1)
output_rate_hist=output.reset_index().fillna(0).drop(columns=['sort','index'])



output_rate_hist['总计']='总计'
col=list(output_rate_hist.columns.values)
output_rate_hist.columns=col
col1=list(output_rate_hist.columns.values)
col1.remove('总计')
tt=output_rate_hist.groupby('总计')[col1].sum().reset_index()
tt=tt.rename({'总计':'使用率'},axis=1)
output_rate_hist=output_rate_hist.append(tt).drop(columns='总计')[col1]

output_rate=output_rate_hist.copy()
col2=list(output_rate.columns.values)
dic2=dict(output_rate[col2].max())

for value in col2:
    try:
        dic2[value]>0
        num=dic2[value]
        output_rate[value]=output_rate[value]/num
        output_rate[value]=output_rate.apply(lambda x: "%.2f%%" % (x[value] * 100) ,axis=1)
    except:
        print(value)
        
        
pct=output_rate.set_index(['使用率']).T
amt=output_rate_hist.set_index(['使用率']).T

pct['统计类型']='百分比'
amt['统计类型']='账户数'

output_amount_total=pd.concat([pct,amt])
output_rate_hist=output_amount_total.reset_index().set_index(['index','统计类型']).sort_index().T
output_rate_hist.columns.names=['业务类型','统计类型']


now=dt.datetime.now()
d=now.strftime("%Y_%m_%d")
#d='2020_07_01'
t=(now + relativedelta(weekday=SA(-1)))
t=pd.to_datetime(t.strftime("%Y/%m/%d"))
#t=pd.to_datetime('2020-07-01')
t_md=t.strftime("%m%d")
d_md=now.strftime("%m%d")
#d_md='0724'


dd=data1[data1['授信时间']>=t].groupby(['授信金额_f','类型'])['寺库用户ID'].count().reset_index().rename({'寺库用户ID':'数量'},axis=1)
tt=dd.pivot(columns='类型' , values='数量' , index='授信金额_f')
mm=tt.reset_index()
mm['sort']= mm['授信金额_f'].apply(lambda x: int(re.findall('^\d+',x)[0]))
output=mm.sort_values(by='sort',ascending=True)
output=output.rename({'授信金额_f':'授信金额'},axis=1)
output_amount_re=output.reset_index().fillna(0).drop(columns=['sort','index'])


output_amount_re['总计']='总计'
col=list(output_amount_re.columns.values)
output_amount_re.columns=col
col1=list(output_amount_re.columns.values)
col1.remove('总计')
tt=output_amount_re.groupby('总计')[col1].sum().reset_index()
tt=tt.rename({'总计':'授信金额'},axis=1)
output_amount_re=output_amount_re.append(tt).drop(columns='总计')[col1]

dd=data1[data1['授信时间']>=t].groupby(['使用率_f','类型'])['寺库用户ID'].count().reset_index().rename({'寺库用户ID':'数量'},axis=1)
tt=dd.pivot(columns='类型' , values='数量' , index='使用率_f')
mm=tt.reset_index()
mm['sort']= mm['使用率_f'].apply(lambda x: int(re.findall('^\d+',x)[0]))
output=mm.sort_values(by='sort',ascending=True)
output=output.rename({'使用率_f':'使用率'},axis=1)
output_rate_re=output.reset_index().fillna(0).drop(columns=['sort','index'])



output_rate_re['总计']='总计'
col=list(output_rate_re.columns.values)
output_rate_re.columns=col
col1=list(output_rate_re.columns.values)
col1.remove('总计')
tt=output_rate_re.groupby('总计')[col1].sum().reset_index()
tt=tt.rename({'总计':'使用率'},axis=1)
output_rate_re=output_rate_re.append(tt).drop(columns='总计')[col1]

d_rate_hist=['历史存量客户用信分布情况']
dd_rate_hist=pd.DataFrame(d_rate_hist)

d_amount_hist=['历史存量客户授信金额分布情况']
dd_amount_hist=pd.DataFrame(d_amount_hist)


d_rate_re=['本周存量客户用信分布情况(账户数)']
dd_rate_re=pd.DataFrame(d_rate_re)

d_amount_re=['本周存量客户授信金额分布情况（账户数）']
dd_amount_re=pd.DataFrame(d_amount_re)


t_mark=['本周通过率情况']
t_mark_df=pd.DataFrame(t_mark)

ot_mark=['逾期情况分析']
ot_mark_df=pd.DataFrame(ot_mark)


til=pd.DataFrame(columns=['类型' ,'类型a' , '申请用户数','通过用户数','通过率' ,'统计时间'])
til['类型a']=[
            #'一键激活',
            '联合授信-新客',
            '联合授信-提额'
            ]

til['类型']=[
            #'一键激活',
            '增量',
            '提额'
            ]
gg=data1[data1['授信时间']>=t].groupby('类型')['授信金额','已用额度'].sum().reset_index().rename({'授信金额':'总授信额度'},axis=1)
gg['额度使用率']=gg['已用额度']/gg['总授信额度']
gg['额度使用率']=gg.apply(lambda x: "%.2f%%" % (x['额度使用率'] * 100) ,axis=1)
til=pd.merge(til,gg,how='left',on='类型')
til1=til[['类型a' , '申请用户数','通过用户数','通过率',	'总授信额度' ,'已用额度' ,'额度使用率' ,'统计时间']]
til1['统计时间']='%s'%(t_md)+'-'+'%s' %(d_md)
til1=til1.rename({'类型a':'类型'},axis=1).fillna(0)



tt=data1.groupby(['类型','逾期天数_f'])['已用额度'].agg(['sum', 'count'])
ddd=tt.pivot_table(columns='逾期天数_f',values=['sum', 'count' ] , index=['类型'])
dd1=ddd['sum'].fillna(0)
dd1['逾期率']=(dd1['0-30']+dd1['30-90']+dd1['90+'])/(dd1['0-30']+dd1['30-90']+dd1['90+']+dd1['0'])
dd1['type']='逾期金额'
# tt=dd.T
# mm=tt.reset_index().rename({'逾期天数_f':'逾期天数'},axis=1)
# mm.columns=['逾期天数','一键激活','增量','提额'] 
dd=ddd['count'].fillna(0)
dd['逾期率']=(dd['0-30']+dd['30-90']+dd['90+'])/(dd['0-30']+dd['30-90']+dd['90+']+dd['0'])
dd['type']='逾期账户数'



# ttt=data1.groupby(['类型','逾期天数_f'])['累计消费'].agg(['sum'])
# dddd=ttt.pivot_table(columns='逾期天数_f',values=['sum' ] , index=['类型'])
# ddd1=dddd['sum'].fillna(0)
# ddd1['逾期率']=(ddd1['0-30']+ddd1['30-90']+ddd1['90+'])/(ddd1['0-30']+ddd1['30-90']+ddd1['90+']+ddd1['0'])
# ddd1['type']='累计交易金额'
# tt=dd.T
# mm1=tt.reset_index().rename({'逾期天数_f':'逾期天数'},axis=1)
# mm1.columns=['逾期天数','一键激活','增量','提额'] 
ot=pd.concat([dd,dd1])
ttt=data1.groupby(['类型'])['累计消费'].agg(['sum'])
leiji=pd.merge(ot,ttt,how='left',on='类型')
leiji['累计交易金额逾期率']=(leiji['0-30']+leiji['30-90']+leiji['90+'])/(leiji['sum'])
leiji['累计交易金额逾期率'][leiji['type']=='逾期账户数']=np.nan
leiji=leiji.drop(columns=['sum'])
ot=leiji.reset_index()
ot['逾期率']=ot.apply(lambda x: "%.2f%%" % (x['逾期率'] * 100) ,axis=1)
ou=ot.set_index(['类型','type']).sort_index().T
ou.index.name='逾期天数'
ou.columns.names=['业务类型','逾期类型']
# ou=pd.concat([mm,mm1],keys=['逾期金额','逾期账户数'],axis=1)

write=pd.ExcelWriter(path+'\\自授信周报_%s.xlsx' %(time))   ##输出结果路径

output_rate_hist.to_excel(write, startcol=2 , startrow=43)
dd_rate_hist.to_excel(write,index=False , startcol=2 , startrow=42,header=False)

output_amount_hist.to_excel(write, startcol=2 , startrow=25)
dd_rate_hist.to_excel(write,index=False , startcol=2 , startrow=24,header=False)



output_rate_re.to_excel(write,index=False , startcol=6 , startrow=10)
dd_rate_re.to_excel(write,index=False , startcol=6 , startrow=9,header=False)

output_amount_re.to_excel(write,index=False , startcol=2 , startrow=10)
dd_rate_re.to_excel(write,index=False , startcol=2 , startrow=9,header=False)

til1.to_excel(write,index=False , startcol=2 , startrow=2)
t_mark_df.to_excel(write,index=False , startcol=2 , startrow=1,header=False)

ou.to_excel(write , startcol=2 , startrow=61)
ot_mark_df.to_excel(write,index=False , startcol=2 , startrow=60,header=False)

write.save()

print('报告已完成')

 


