import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.express as px

import io

st.set_page_config(page_title="数据发布工具集", layout="wide")

s = st.sidebar.radio('数据发布工具集',('数据发布统计','数据管理全景跟踪'))

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
                    q = '(%s == "G") or (%s == "N") and (not (%s == ""))' % (s, s, s)

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
     
                    #lastGCount = 0
                    #lastGDate = ''
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
                        plan = {'总数':0, '计划':'', '完成':0, '未完成':0, '未到期':0, '计划完成率':0, '实际完成率':0}
                        if vDate in plans:
                            plan = plans[vDate]
                        #pSum = plan['总数'] + vCnt
                        #pGSum = plan['完成'] + gCnt
                        #pNSum = plan['未完成'] + nCnt

                        #nCount = int(list(xdf.query('(%s > "%s") and (%s == "N")' % (c, dtsNow, s)).agg('sum', numeric_only=True).values)[4])

                        gCount = int(list(xdf.query('(%s <= "%s") and (%s == "G")' % (c, vDate, s)).agg('sum', numeric_only=True).values)[4])

                        plan['总数'] = dSum
                        plan['完成'] = gCount 
                        
                        #if not (lastGDate == vDate):
                        #    if gCount == lastGCount:
                        #        plan['完成'] = ''
                        #    else:
                        #        lastGCount = gCount 
                        #        lastGDate = vDate                        

                        npCount = int(list(xdf.query('(%s > "%s")' % (c, vDate)).agg('sum', numeric_only=True).values)[4])

                        plan['未到期'] = npCount
                        plan['计划'] = int(dSum - npCount)
                        if vDate > dtsNow:
                            #plan['计划2'] = plan['计划']
                            #plan['完成'] = ''
                            plan['未完成'] = ''

                        plan['未完成'] = int(dSum - npCount - gCount)

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

def showPubDataTool():
    pubObjs = []
    st.header('数据管理全景跟踪')
    uploaded_pubFile = st.file_uploader("上传数据管理全景跟踪表", type=["xlsx"])
    if uploaded_pubFile is not None:
        fname = uploaded_pubFile.name
        st.success('数据管理全景跟踪表:%s已上传完成' % fname)     
        df = pd.read_excel(uploaded_pubFile, header=1)
        rows = list(df.values)
        pubObj = None
        opClass = None
        opName = None
        opIdx = None
        opType = None
        for row in rows:
            idx = row[0]
            pClass = row[1]
            pName = row[2]
            pType = row[3]
            pTask = row[5]
            if pClass == pClass:
                opClass = pClass
            if pName == pName:
                opName = pName
            if idx == idx:
                opIdx = idx                
            if pType == pType:
                opType = pType
            colName = row[4]
            colValues = row[5:16]
            colValues = [str(x) for x in colValues]
            problems = row[16]   
            respers = row[17]
            resons = row[18]
            if pName == pName: #车型行
                pubObj = {'序号':opIdx, '产品线':opClass, '项目':opName, '立项':opType, '任务':pTask, '问题描述':problems, '责任部门/人':respers, '原因分析':resons} 
                pubObjs.append(pubObj)
                #'车型行: idx, pClass, pName, pType, colName, problems, respers, resons'  
                #opIdx, opClass, opName, opType, problems, respers, resons   
                #'pubObj'
                #pubObj            
            else:  #数据行
                #'数据行: idx, pClass, pName, pType, colName'  
                if pubObj:
                    if colName in ['完成率', '延期天数']:
                        pass
                    else:
                        if colName in ['计划完成时间', '实际完成时间']:
                            cvs = []
                            for cv in colValues:
                                if cv in ['','/','nan']:
                                    cvs.append('')
                                else:
                                    d = datetime.strptime(cv, '%Y-%m-%d %H:%M:%S')
                                    ds = datetime.strftime(d, '%Y-%m-%d')
                                    cvs.append(ds)
                            pubObj[colName] = cvs
                        elif colName in ['计划数量', '完成数量']:
                            cvs = []
                            for cv in colValues:
                                cvs.append(int(cv))
                            pubObj[colName] = cvs
                        else:
                            pubObj[colName] = colValues 
        pubRecords = []
        for po in pubObjs:
            pdps = po['部门']
            pts = po['计划数量']
            cts = po['完成数量']
            vs = []
            for i in range(len(pts)):
                ppt = pts[i]
                pct = cts[i]
                v = 0
                if ppt > 0:
                    v = round(pct / ppt, 2)
                    v = '%d%%' % (v*100)
                vs.append(v)
            po['完成率'] = vs
            pcvs = vs
            pds = po['计划完成时间']
            cds = po['实际完成时间']
            vs = []
            for i in range(len(pds)):
                ppd = pds[i]
                pcd = cds[i]
                v = 0
                if ppd:
                    p_t = datetime.strptime(ppd, '%Y-%m-%d')
                    if pcd:
                        d_t = datetime.strptime(pcd, '%Y-%m-%d') 
                    else:
                        d_t = datetime.now()   
                    v = max((d_t - p_t).days, 0)
                vs.append(v)
            po['延期天数'] = vs
            dts = vs

            pRecs = []
            for i in range(len(pdps)):
                rDept = pdps[i]
                rPt = pts[i]
                rCt = cts[i]
                rCv = pcvs[i]
                rPd = pds[i]
                rCd = cds[i]
                rDt = dts[i]
                pc = po['产品线']
                pn = po['项目']
                pn = '%s-%s' % (pc, pn)
                pRecs.append({'部门':rDept,'计划数量':rPt,'完成数量':rCt,'完成率':rCv, 
                                '计划完成时间':rPd,'实际完成时间':rCd,'延期天数':rDt})
                pubRec1 = {'产品线':pc,'项目':pn,'立项':po['立项'],
                            '部门':rDept,'计划数量':rPt,'数量':rCt,'完成率':rCv,
                            '计划完成时间':rPd,'实际完成时间':rCd,'延期天数':rDt,'完成状态':'已完成'}
                pubRec2 = {'产品线':pc,'项目':pn,'立项':po['立项'],
                            '部门':rDept,'计划数量':rPt,'数量':rPt-rCt,'完成率':rCv,
                            '计划完成时间':rPd,'实际完成时间':rCd,'延期天数':rDt,'完成状态':'未完成'}
                pubRecords.append(pubRec1)
                pubRecords.append(pubRec2)
                
            po['callRecords'] = pRecs

        #st.dataframe(pubObjs, use_container_width=True)
        #st.dataframe(pubRecords)
        
        if 0:
            pcs = st.selectbox('选择产品线', ['T1N', 'T1J', 'T1K', 'T1P', 'T1L', 'T-2'])
            tdf = pd.DataFrame(pubObjs)
            tdf = tdf.query('pClass == "%s"' % pcs)
            rcs = tdf.to_dict(orient='records')            
            for rc in rcs:
                p_cls = rc['pClass']
                p_prj = rc['pName']
                p_typ = rc['pType']
                p_tsk = rc['pTask']
                p_prb = rc['problems']
                p_rps = rc['respers']
                p_rss = rc['resons']
                p_dps = rc['部门']
                p_pns = rc['计划数量']
                p_cns = rc['完成数量']
                p_pts = rc['计划完成时间']
                p_cts = rc['实际完成时间']
                st.divider()
                '项目：%s - %s - %s ' % (p_cls, p_prj, p_typ)
                '任务：%s' % p_tsk
                '问题：%s' % p_prb
                '责任人：%s' % p_rps
                '原因：%s' % p_rss
                drs = []
                totalPn = 0
                totalCn = 0
                for i in range(len(p_dps)):
                    t_dp = p_dps[i]
                    t_pn = int(p_pns[i])
                    t_cn = int(p_cns[i])
                    t_pt = p_pts[i]
                    t_ct = p_cts[i]
                    if t_dp == '总计':
                        t_pn = totalPn
                        t_cn = totalCn
                    else:
                        totalPn += t_pn
                        totalCn += t_cn
                    t_cv = 0
                    if t_pn > 0:
                        t_cv  = t_cn / t_pn
                    if t_ct in ['', '/', 'nan']:
                        t_ct = ''
                    else:
                        t_ct = datetime.strptime(t_ct, '%Y-%m-%d %H:%M:%S')
                    if t_pt not in ['', '/', 'nan']:
                        t_pt = datetime.strptime(t_pt, '%Y-%m-%d %H:%M:%S')
                    else:
                        t_pt = ''
                    t_dt = 0
                    if t_pt:
                        if t_ct:
                            t_dt = max((t_ct - t_pt).days, 0) 
                        else:
                            t_dt = max((datetime.now() - t_pt).days, 0)                                       
                    dr = {'部门':t_dp, '计划数量':t_pn, '完成数量':t_cn, '完成率':t_cv, '计划完成时间':t_pt, '实际完成时间':t_ct, '延期天数':t_dt}
                    drs.append(dr)
                st.dataframe(drs, use_container_width=True)
            st.divider()
        
        deptRows = []
        projRows = []

        if pubObjs:  
            po = pubObjs[0]
            depts = po['部门']
            dn = len(depts)            
            projRows.append(['部门数据发布全景'] + ['']*(dn+7))
            projRows.append(['序号'] + ['项目','项目' ] + ['立项'] + ['数据发布']+['']*(dn) + ['问题描述','责任部门/人','原因分析'])
            for pubObj in pubObjs:    
                pIdx = int(pubObj['序号'])                
                pClass = pubObj['产品线']
                pName = pubObj['项目']
                pType = pubObj['立项']
                pTask = pubObj['任务']

                pPrb = pubObj['问题描述']
                pRps = pubObj['责任部门/人']
                pRrs = pubObj['原因分析']
                pDps = pubObj['部门']
                pPns = pubObj['计划数量']
                pCns = pubObj['完成数量']
                pCvs = pubObj['完成率']
                pPts = pubObj['计划完成时间']
                pCts = pubObj['实际完成时间']
                pDts = pubObj['延期天数']

                projRows.append([pIdx, pClass, pName, pType, '任务'] + [pTask]*dn+[pPrb, pRps, pRrs])  
                projRows.append([pIdx, pClass, pName, pType, '部门'] + pDps +['']*3) 
                projRows.append([pIdx, pClass, pName, pType, '计划数量'] + pPns +['']*3)
                projRows.append([pIdx, pClass, pName, pType, '完成数量'] + pCns +['']*3) 
                projRows.append([pIdx, pClass, pName, pType, '完成率'] + pCvs +['']*3)
                projRows.append([pIdx, pClass, pName, pType, '计划完成时间'] + pPts +['']*3)
                projRows.append([pIdx, pClass, pName, pType, '实际完成时间'] + pCts +['']*3)
                projRows.append([pIdx, pClass, pName, pType, '延期天数'] + pDts +['']*3)
            df = pd.DataFrame(projRows)
            def applyColor(dtf):
                ss = dtf.values
                n = len(dtf) 
                cs = ['background-color: #ddd']*2 + [''] * (n-2)
                pn = int((n-2)/8)
                pTask = str(ss[2]).lower()
                pCol = ss[1]                   
                for i in range(pn):         
                    if pTask == '任务':
                        for j in range(8):
                            cs[i*8+2+j] = 'background-color: #FFE'   
                    elif pTask in ['m0', 'm1', 'm2']:
                        cs[i*8+3] = 'background-color: #FFE'
                        pPn = int(ss[i*8+4])
                        pCn = int(ss[i*8+5])
                        pCv = ss[i*8+6]
                        pPt = ss[i*8+7]
                        pCt = ss[i*8+8]
                        pDt = ss[i*8+9]
                        cs[i*8+2] = 'background-color: #BFF'
                        color = '#DDD'
                        if (pPn == pCn):
                            color = '#9f9'  
                        elif pDt:
                            color = '#f99'    
                        if color:
                            for j in range(3):
                                cs[i*8+4+j] = 'background-color: %s' % color
                    else:
                        cs[i*8+2] = ''
                    if pCol in ['序号','项目','立项']: 
                        for j in range(8):
                            cs[i*8+2+j] = 'background-color: #BFF'
                return cs
            deptDf = df.style.apply(applyColor, axis=0)
            st.dataframe(deptDf, use_container_width=True)                 

            rdf = pd.DataFrame(pubRecords)            
            tdf = rdf.query('部门 != "总计"').groupby(by=['部门','完成状态'], as_index=False).sum(numeric_only=True)
            fig = px.bar(tdf, x='部门', y='数量', title='部门发布情况统计', text='数量', height=400, color='完成状态')
            fig.update_layout(xaxis={'categoryarray':['动总','底盘','电控','外饰','内饰','车身','电器','网联','热管理','架构'],'categoryorder':'array'})
            st.plotly_chart(fig, use_container_width=True)

        if pubObjs:                
            deptRows.append(['部门','项目'])
            deptRows.append(['',''])
            deptRows.append(['','立项'])
            deptRows.append(['','任务'])
            #lClass = ''
            for pubObj in pubObjs:
                pClass = pubObj['产品线']
                #if pClass == lClass:
                #    deptRows[0].append('') 
                #else:
                #    deptRows[0].append(pClass)
                #    lClass = pClass 
                deptRows[0].append(pClass)
                deptRows[1].append(pubObj['项目']) 
                deptRows[2].append(pubObj['立项'])   
                deptRows[3].append(pubObj['任务']) 
            deptRows[0].append('总计')   
            deptRows[1].append('')   
            deptRows[2].append('')   
            deptRows[3].append('')  
            pDepts = pubObjs[0]['部门']    
            cols = ['计划数量', '完成数量', '完成率', '计划完成时间', '实际完成时间', '延期天数']
            deptIdx = 0
            p_pns = [0] * len(pubObjs) 
            p_cns = [0] * len(pubObjs) 
            dpPts = [None] * len(pubObjs) 
            dpCts = [None] * len(pubObjs) 
            for pDept in pDepts:
                dpPns = 0
                dpCns = 0                    
                for col in cols:
                    tRow = [pDept, col]                                                                       
                    for i in range(len(pubObjs)):
                        pubObj = pubObjs[i]
                        v = pubObj[col][deptIdx]
                        if pDept == '总计':
                            if col == '计划数量':
                                tRow.append(p_pns[i])
                                dpPns += int(v) 
                            elif col == '完成数量':
                                tRow.append(p_cns[i])
                                dpCns += int(v)
                            elif col == '完成率':
                                pn = p_pns[i]
                                cn = p_cns[i]
                                v = 0
                                if not(pn == 0):
                                    v = round(cn / pn, 2)
                                tRow.append('%d%%' % (v*100))
                        else:
                            if col == '计划数量':
                                p_pns[i] = p_pns[i] + int(v)
                                dpPns += int(v) 
                            elif col == '完成数量':
                                p_cns[i] = p_cns[i] + int(v)
                                dpCns += int(v)
                            elif col == '完成率':
                                pn = int(pubObj['计划数量'][deptIdx])
                                cn = int(pubObj['完成数量'][deptIdx])
                                v = '/'
                                if not(pn == 0):
                                    v = round(cn / pn, 4)
                                    v = '%d%%' % (v*100)
                            tRow.append(v)
                    if col == '计划数量':
                        tRow.append(dpPns) 
                    elif col == '完成数量':
                        tRow.append(dpCns) 
                    elif col == '完成率':
                        v = 0
                        if not(dpPns == 0):
                            v = round(dpCns / dpPns, 2)
                        tRow.append('%d%%' % (v*100))
                    else:
                        tRow.append('')    
                    if (pDept == '总计') and (col in ['计划完成时间', '实际完成时间', '延期天数']):
                        pass
                    else:
                        deptRows.append(tRow)                    
                deptIdx += 1
        df = pd.DataFrame(deptRows)
        def applyColor(dtf):
            ss = dtf.values
            task = str(ss[3]).lower()
            dpn = int((len(dtf) - 7)/6)
            cs = [''] * (len(dtf)-3) + ['background-color: #CCB']*3
            headColor = '#BFF'
            if ss[0] == '总计':
                cs = ['background-color: #CCB'] * len(dtf)
            elif ss[0] in ['部门', '项目']:
                cs = ['background-color: #EEE'] * 4 + ['background-color: #FFE'] * (len(dtf) - 4 - 3) + ['background-color: #CCB']*3
            else:
                for i in range(4):
                    cs[i] = 'background-color: %s' % headColor
                if task in ['m0', 'm1', 'm2']:
                    pState = 0
                    pStateColor = '#DDD'                        
                    for i in range(dpn):
                        n = i*6+4
                        pn,cn,cv,pt,ct,dt = ss[n:n+6]
                        #pn,cn,cv,pt,ct,dt
                        color = '#DDD'
                        if (pn == cn):
                            pState = max(pState, 0)
                            color = '#9f9'
                            pStateColor = color    
                        elif dt:
                            pState = max(pState, 2)
                            color = '#f99' 
                        else:
                            pState = max(pState, 1)      
                        if color:
                            for i in range(3):
                                cs[n+i] = 'background-color: %s' % color 
                    if pState == 0:
                        pStateColor = '#9f9'
                    elif pState == 1:
                        pStateColor = '#ddd'
                    elif pState == 2:
                        pStateColor = '#f99'
                    cs[1] = 'background-color: %s' % pStateColor 
            return cs
        projDf = df.style.apply(applyColor, axis=0)
        st.dataframe(projDf, use_container_width=True)

        rdf = pd.DataFrame(pubRecords)
        tdf = rdf.query('部门 != "总计"').groupby(by=['项目','完成状态'], as_index=False).sum(numeric_only=True)
        fig = px.bar(tdf, x='项目', y='数量', title='项目发布情况统计', text='数量', height=400, color='完成状态')
        st.plotly_chart(fig, use_container_width=True)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            sheet_dept = '部门全景跟踪'
            sheet_proj = '项目全景跟踪'
            deptDf.to_excel(writer, sheet_name=sheet_dept, index=False) 
            projDf.to_excel(writer, sheet_name=sheet_proj, index=False)                  
            writer.close()
            fileName= '数据发布全景跟踪表_%s.xlsx' % (datetime.now())
            st.write(fileName)
            st.download_button(
                label="导出全景跟踪表",
                data=buffer,
                file_name=fileName,
                mime='application/vnd.ms-excel')
            
if s == '数据发布统计':
    ShowDataWorkStatistics()
elif s == '数据管理全景跟踪':
    showPubDataTool()
else:
    st.title('其他数据发布工具持续开发中，敬请期待！')

