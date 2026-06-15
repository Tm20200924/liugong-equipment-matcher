# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import os, sys, json, io
from datetime import datetime
import openpyxl

st.set_page_config(page_title='LiuGong Matcher v19', page_icon='🔧', layout='wide')

# ==================== ENGINE LOADING ====================

@st.cache_resource
def load_engine():
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'engine'))
    import engine_v19 as eng
    return eng

@st.cache_data(ttl=3600)
def load_data():
    base = os.path.dirname(__file__)
    eng.CONFIG_PATH = os.path.join(base, 'engine', 'config.json')
    eng.DB_PATH = os.path.join(base, 'engine', 'product_db_full.json')
    eng.COMP_DB_PATH = os.path.join(base, 'engine', 'competitor_db.json')
    config = eng.load_config()
    db = eng.load_db()
    return config, db

eng = load_engine()
config, db = load_data()

# ==================== SIDEBAR ====================

with st.sidebar:
    st.header('⚙️ 设置')
    st.subheader('汇率与税率')
    rate = st.number_input('汇率 (CNY→RUB)', value=config.get('exchange_rate_rub_cny', 11.5), step=0.1, format='%.1f')
    duty = st.number_input('关税 (%)', value=config.get('customs_duty_rate_pct', 5.0), step=0.1, format='%.1f')
    proc = st.number_input('报关手续费 (%)', value=config.get('customs_processing_rate_pct', 0.3), step=0.1, format='%.1f')
    vat = st.number_input('增值税 (%)', value=config.get('vat_rate_pct', 22.0), step=0.1, format='%.1f')
    warehouse = st.number_input('仓储费 (RUB)', value=config.get('customs_warehouse_fee_rub', 20000), step=1000)
    customs_fee = st.number_input('海关费 (RUB)', value=config.get('customs_fee_rub', 6000), step=1000)
    agent_fee = st.number_input('代理费 (RUB)', value=config.get('customs_agent_fee_rub', 30000), step=1000)
    
    st.divider()
    st.caption(f'产品库: {len(db)} 个型号 | v19')

# Update config with UI values
config['exchange_rate_rub_cny'] = rate
config['customs_duty_rate_pct'] = duty
config['customs_processing_rate_pct'] = proc
config['vat_rate_pct'] = vat
config['customs_warehouse_fee_rub'] = warehouse
config['customs_fee_rub'] = customs_fee
config['customs_agent_fee_rub'] = agent_fee

# ==================== HEADER ====================

st.title('🔧 LiuGong 柳工设备匹配引擎 v19')
st.caption('上传询盘 → 精准匹配 → DAP满洲里到岸价计算 → 交叉验证')

# ==================== TAB 1: INPUT ====================

tab1, tab2, tab3 = st.tabs(['📥 询盘输入', '📊 匹配结果', '✅ 验证报告'])

with tab1:
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader('上传文件')
        uploaded = st.file_uploader('上传询盘文件 (Excel/CSV)', type=['xlsx', 'xls', 'csv'],
                                    help='支持多行询盘，引擎会自动识别设备名称、数量和需求')
    
    with col2:
        st.subheader('或手动输入')
        manual_text = st.text_area('手动输入询盘（每行一条）', height=150,
                                   placeholder='示例:\n挖掘机 20吨 2台\n压路机 3吨级 6台\n推土机бульдозер TY230 4台')
    
    if st.button('🚀 开始匹配', type='primary', use_container_width=True):
        with st.spinner('正在匹配...'):
            inquiries = []
            
            if uploaded:
                # Save to temp and parse
                tmp_path = os.path.join(os.path.dirname(__file__), 'work', f'_upload_{datetime.now().strftime("%Y%m%d%H%M%S")}.xlsx')
                os.makedirs(os.path.dirname(tmp_path), exist_ok=True)
                with open(tmp_path, 'wb') as f:
                    f.write(uploaded.getbuffer())
                inquiries = eng.parse_inquiry_excel(tmp_path)
                st.success(f'从文件解析 {len(inquiries)} 条询盘')
            
            if manual_text.strip():
                lines = [l.strip() for l in manual_text.strip().split('\n') if l.strip()]
                for i, line in enumerate(lines):
                    # Try to extract qty
                    qty = 1
                    import re
                    qm = re.search(r'(\d+)\s*[台套辆]', line)
                    if qm: qty = int(qm.group(1))
                    inquiries.append({'seq': str(i+1), 'name': line, 'qty': qty})
                st.success(f'从文本解析 {len(inquiries)} 条询盘')
            
            if inquiries:
                st.session_state['inquiries'] = inquiries
                st.session_state['results'] = []
                for item in inquiries:
                    cands, ton, bucket, ver = eng.match_products(item['name'], db, config)
                    st.session_state['results'].append({
                        'item': item, 'candidates': cands, 'ton': ton,
                        'bucket': bucket, 'verification': ver
                    })
            else:
                st.warning('请上传文件或输入询盘内容')

# ==================== TAB 2: RESULTS ====================

with tab2:
    if 'results' not in st.session_state or not st.session_state['results']:
        st.info('👈 请先在"询盘输入"标签页上传文件或输入询盘，然后点击"开始匹配"')
    else:
        results = st.session_state['results']
        
        # Summary stats
        matched = sum(1 for r in results if r['candidates'])
        st.metric('匹配统计', f'{matched}/{len(results)} 条匹配成功')
        
        # Build display table
        rows = []
        for r in results:
            item = r['item']
            cands = r['candidates']
            cat = eng.classify_inquiry(item['name']) or ''
            comp_brand, _ = eng.detect_competitor(item['name'])
            if comp_brand: cat = f'竞品({comp_brand})' if not cat else f'{cat}/竞品({comp_brand})'
            
            if not cands:
                rows.append({
                    '序号': item['seq'], '询盘需求': item['name'], '数量': item['qty'],
                    '设备类别': cat or '非设备', '推荐型号': '—', '匹配度': '—',
                    'DAP总价(RUB)': '—', '推荐理由': '非柳工产品' if eng.is_non_liugong(item['name']) else '未找到匹配',
                    '验证等级': '—'
                })
            else:
                for ci, cand in enumerate(cands[:3]):
                    p = cand['product']
                    cost = eng.calc_dap(p.get('dap_price_cny', 0) or 0, p.get('scrap_tax_rub', 0) or 0, config)
                    reasons = '; '.join(cand['reasons']) if cand['reasons'] else '综合匹配'
                    ver_level = ''
                    if r['verification'] and r['verification'].get('top_verification'):
                        ver_level = r['verification']['top_verification'].get('level', '')
                    rows.append({
                        '序号': item['seq'] if ci == 0 else '',
                        '询盘需求': item['name'] if ci == 0 else '',
                        '数量': item['qty'] if ci == 0 else '',
                        '设备类别': cat if ci == 0 else '',
                        '推荐型号': p['model'],
                        '匹配度': f'{cand["score"]}分' + (' ⭐推荐' if ci == 0 else f' 备选{ci}'),
                        'DAP总价(RUB)': f'{cost["total_rub"]:,.0f}',
                        '推荐理由': reasons,
                        '验证等级': ver_level
                    })
        
        df = pd.DataFrame(rows)
        st.dataframe(df, use_container_width=True, height=600,
                     column_config={
                         '推荐型号': st.column_config.TextColumn(width='small'),
                         '匹配度': st.column_config.TextColumn(width='small'),
                         'DAP总价(RUB)': st.column_config.TextColumn(width='medium'),
                     })
        
        # Download buttons
        col1, col2 = st.columns(2)
        with col1:
            # Download Excel report
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='匹配结果')
            st.download_button('📥 下载Excel报告', buf.getvalue(),
                               f'Report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx',
                               'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        with col2:
            csv = df.to_csv(index=False).encode('utf-8-sig')
            st.download_button('📥 下载CSV', csv,
                               f'Report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv',
                               'text/csv')

# ==================== TAB 3: VERIFICATION ====================

with tab3:
    if 'results' not in st.session_state or not st.session_state['results']:
        st.info('👈 请先完成匹配')
    else:
        results = st.session_state['results']
        
        ver_rows = []
        for r in results:
            item = r['item']
            ver = r.get('verification', {})
            if not ver: continue
            
            tv = ver.get('top_verification', {}) or {}
            ver_rows.append({
                '询盘': item['name'][:40],
                '分类': ver.get('classified_cat', ''),
                '提取吨位': ver.get('extracted_ton', 0),
                '提取马力': ver.get('extracted_hp', 0),
                '竞品': ver.get('competitor', ''),
                '本地吨位': tv.get('local_ton', 0),
                '联网吨位': tv.get('online_ton', 0),
                '本地马力': tv.get('local_hp', 0),
                '联网马力': tv.get('online_hp', 0),
                '验证等级': tv.get('level', ''),
                '状态': tv.get('status', ''),
            })
        
        if ver_rows:
            st.subheader('交叉验证详情')
            ver_df = pd.DataFrame(ver_rows)
            st.dataframe(ver_df, use_container_width=True)
            
            # Confidence summary
            dual = sum(1 for v in ver_rows if v['验证等级'] == 'dual')
            single = sum(1 for v in ver_rows if v['验证等级'] == 'single')
            conflict = sum(1 for v in ver_rows if v['验证等级'] == 'conflict')
            
            col1, col2, col3, col4 = st.columns(4)
            col1.metric('双重验证', dual)
            col2.metric('单方验证', single)
            col3.metric('数据冲突', conflict, delta_color='inverse')
            col4.metric('总计', len(ver_rows))
        else:
            st.info('暂无验证数据')

# ==================== FOOTER ====================

st.divider()
st.caption(f'LiuGong Equipment Matcher v19 | DAP满洲里到岸价 | 产品库: {len(db)} 个型号 | {datetime.now().strftime("%Y-%m-%d %H:%M")}')
