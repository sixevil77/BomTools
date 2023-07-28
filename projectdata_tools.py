import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="数据发布工具集", layout="wide")

s = st.sidebar.radio('数据发布工具集',('数据发布统计',))

from datetime import datetime

def Bar2(container, data, x, y, text,  title='柱状图',ori='v',height=300):
    tab1, tab2 = container.tabs(['图表','数据'])    
    fig = px.bar(data, x=x, y=y, title=title, text=text, orientation=ori, height=height)
    if ori == 'v':
        fig = px.bar(data, x=x, y=y, title=title, text=text, orientation=ori, height=height)
    #fig.update_layout(xaxis_title=xTitle,yaxis_title=yTitle)
    #fig.update_traces(textposition="outside")
    fig.update_layout(
        yaxis={
            'tickformat': '.2f'
        }
    )
    tab1.plotly_chart(fig, use_container_width=True)
    tab2.dataframe(data, use_container_width=True)

def ShowDataWorkStatistics():
    dtsNow = datetime.strftime(datetime.now(), '%Y-%m-%d %H:%M:%S')
    uploaded_file = st.file_uploader("请上传需要进行统计的数据表", type=["xlsx"])
    if uploaded_file is not None:      
        try:
            df = pd.read_excel(uploaded_file)

            cs = ['M0数据发布计划', 'M1数据发布计划', 'M2数据发布计划']
            ss = ['M0完成状态', 'M1完成状态', 'M2完成状态']
            xqs = {'高+低':'(续航 == "高") or (续航 == "低")', '高':'续航 == "高"', '低':'续航 == "低"'}
            
            
                #st.subheader(xTitle)
                #xdf            

            for i in range(0,len(cs)):                
                c = cs[i]
                s = ss[i]
                for xk in xqs:
                    xTitle = '续航:%s' % xk
                    xQuery = xqs[xk]
                    hdf = df.query(xQuery)
                    st.subheader(c+' - 部门统计 - '+xTitle)
                    #q = '(%s != "/") and (%s != "TBD") and (%s == %s)' % (c,c,c,c)
                    #q = 'and (%s == %s) and (%s != "/")' % (c,c,c,c,s)
                    q = '(%s == "G") or (%s == "N")' % (s, s)

                    tdf = hdf.query(q)
                    xdf = tdf.groupby(['部门', s], as_index=False ).agg('count')
                
                    #tdf
                    rs = {}
                    vs = list(xdf.values)

                    for v in vs:
                        dept = v[0]
                        stat = v[1]
                        cnt = v[4]
                        r = {'未完成':0, '已完成':0}
                        if dept in rs:
                            r = rs[dept]
                        if stat == 'G':
                            r['已完成'] = cnt
                        elif stat == 'N':
                            r['未完成'] = cnt
                        rs[dept] = r
                    
                    records = [{'部门':'应完成'},{'部门':'已完成'},{'部门':'未完成'},{'部门':'完成率'}]
                    records1 = []
                    gsum = nsum = dsum = 0
                    depts = ['动总','底盘','电控','内饰','车身','电子电器','智能网联','热管理','外饰']
                    for dept in depts:
                        cnt = gcnt = ncnt = rate = 0
                        if dept in rs:
                            r = rs[dept]
                            gcnt = r['已完成']
                            ncnt = r['未完成']
                            cnt = gcnt + ncnt
                            gsum += gcnt
                            nsum += ncnt
                            dsum += cnt
                            rate = float(gcnt) / float(cnt)

                        records[0][dept] = cnt
                        records[1][dept] = gcnt
                        records[2][dept] = ncnt                        
                        records[3][dept] = '%.2f%%' % (rate*100)

                        r1 = {'部门':dept, '应完成':cnt, '已完成':gcnt, '未完成':ncnt, '完成率':rate}
                        records1.append(r1)

                    records[0]['汇总'] = dsum
                    records[1]['汇总'] = gsum
                    records[2]['汇总'] = nsum                    
                    records[3]['汇总'] = '%.2f%%' % (float(gsum) / float(dsum) * 100)

                    col1, col2 = st.columns([2,1])

                    col1.dataframe(records,use_container_width=True, height=400)                
                    #col1.dataframe(records1,use_container_width=True)

                    tdf1 = pd.DataFrame(records1)
                    #tdf1 = tdf1.sort_values(by='完成率',ascending=False,inplace=False)
                    tdf1['完成率'] = tdf1.apply(lambda col:'%.2f%%' % (col['完成率']*100), axis=1)                
                    #Bar2(col2, tdf, '部门', '完成率', '完成率','部门数据发布完成率', 'v', 400)

                    fig = px.bar(tdf1, x='部门', y='应完成', title='部门数据发布状态', text='应完成', orientation='v', height=400, barmode='group',
                                color_discrete_sequence=["#CC3333"])
                    fig2 = px.bar(tdf1, x='部门', y='已完成', title='title', text='已完成', orientation='v', height=400,
                                color_discrete_sequence=["#3333CC"])
                    fig.add_trace(fig2.data[0])
                    #fig.update_traces(textposition="outside")                
                    col2.plotly_chart(fig, use_container_width=True)

                    #### 续航表格
                    dSum = len(tdf)
                    tdf = tdf.groupby([c, s], as_index=False).agg('count')
                    tdf[c] = tdf.apply(lambda col: str(col[c]), axis=1)
                    xdf = tdf.sort_values(by=c,ascending=True,inplace=False)
                    vs = list(xdf.values)
                    plans = {}
                    gSum = 0                 
                    for v in vs:
                        vDate = v[0]
                        vStat = v[1]
                        vCnt = v[4]
                        gCnt = nCnt = 0
                        if vStat == 'G':
                            gCnt = vCnt
                            gSum += vCnt
                        elif vStat == 'N':
                            nCnt = vCnt
                        plan = {'总数':0, '计划1':'', '计划2':'', '完成':0, '未完成':0, '未到期':0, '计划完成率':0, '实际完成率':0}
                        if vDate in plans:
                            plan = plans[vDate]
                        #pSum = plan['总数'] + vCnt
                        #pGSum = plan['完成'] + gCnt
                        #pNSum = plan['未完成'] + nCnt

                        nCount = int(list(xdf.query('(%s > "%s") and (%s == "N")' % (c, dtsNow, s)).agg('sum', numeric_only=True).values)[4])

                        gCount = int(list(xdf.query('(%s <= "%s") and (%s == "G")' % (c, vDate, s)).agg('sum', numeric_only=True).values)[4])

                        plan['总数'] = dSum
                        plan['完成'] = gCount
                        plan['未完成'] = max((dSum-gCount), nCount)

                        npCount = int(list(xdf.query('(%s > "%s")' % (c, vDate)).agg('sum', numeric_only=True).values)[4])

                        plan['未到期'] = npCount
                        plan['计划1'] = int(dSum - npCount)
                        if vDate > dtsNow:
                            plan['计划2'] = plan['计划1']
                            plan['完成'] = ''
                            plan['未完成'] = ''

                        plan['计划完成率'] = '%.2f%%' % ((dSum-npCount) / dSum * 100)
                        plan['实际完成率'] = '%.2f%%' % (gSum / dSum * 100)
                        plans[vDate] = plan
                    rs = []
                    for d in plans:
                        plan = plans[d]
                        ds = d
                        if ds == 'TBD':
                            pass
                        elif not (ds == '/'):
                            dd = datetime.strptime(d,  '%Y-%m-%d %H:%M:%S')
                            ds = datetime.strftime(dd, '%Y-%m-%d')
                        p = {'日期':ds}
                        for k in plan:
                            p[k] = plan[k]
                        rs.append(p) 
                    st.subheader(c+' - %s' % xTitle)       
                    st.dataframe(rs,use_container_width=True)
                    st.divider()

        except Exception as e:
            e
            
if s == '数据发布统计':
    ShowDataWorkStatistics()
else:
    st.title('其他数据发布工具持续开发中，敬请期待！')

