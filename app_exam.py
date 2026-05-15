import streamlit as st
import pandas as pd
import numpy as np
import io
import pulp
import traceback
import random
import openpyxl
from datetime import datetime

# ==========================================
# 1. 網頁頁面配置
# ==========================================
st.set_page_config(page_title="段考監考終極自動化", page_icon="🏫", layout="wide")
st.title("🏫 試務組-段考監考全自動化系統 (旗艦美學版)")
st.info("💡 最終優化：現在連『監考一覽表』也會完美保留原始 Excel 的框線、粗細與字體格式！")

# --- 初始化狀態 ---
if 'results' not in st.session_state:
    st.session_state['results'] = None
if 'uploader_key' not in st.session_state:
    st.session_state['uploader_key'] = 0

# ==========================================
# 2. 輔助功能定義
# ==========================================
def to_excel_bytes(df, header_df=None):
    output = io.BytesIO()
    if header_df is not None:
        df.columns = header_df.columns
        final_out = pd.concat([header_df, df], ignore_index=True)
    else:
        final_out = df
    final_out = final_out.fillna("")
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        final_out.to_excel(writer, index=False, header=False)
    return output.getvalue()

# 【字串全面淨化器】
def clean_str(s):
    if pd.isna(s) or s is None: return ""
    s = str(s).strip().replace('ㄧ', '一').replace(' ', '').replace('　', '').replace('\n', '').replace('\r', '')
    s = s.translate(str.maketrans('１２３４５６７８９０', '1234567890'))
    return s

def normalize_subject(s):
    s = clean_str(s)
    aliases = {'國文':'國語文', '英文':'英語文', '公社':'公民與社會', '公民':'公民與社會', 
               '地科':'地球科學', '健護':'健康與護理', '護理':'健康與護理', 
               '國防':'全民國防教育', '生科':'生活科技', '應數':'應用數學'}
    return aliases.get(s, s)

# 【超級模糊尋找任課教師】
def get_teacher_fuzzy(cls, subj, course_dict):
    if (cls, subj) in course_dict: return course_dict[(cls, subj)]
    clean_target = subj.replace('選修', '').replace('彈性學習', '').replace('補強', '').replace('-', '')
    for (c, s), t in course_dict.items():
        if c == cls:
            s_clean = s.replace('選修', '').replace('彈性學習', '').replace('補強', '').replace('-', '')
            if clean_target and (clean_target in s_clean or s_clean in clean_target):
                return t
    if len(clean_target) >= 2:
        for (c, s), t in course_dict.items():
            if c == cls and clean_target[:2] in s:
                return t
    return ""

# ==========================================
# 3. 介面佈局
# ==========================================
st.divider()
col1, col2 = st.columns([1, 1], gap="large")

with col1:
    st.subheader("📂 1. 上傳排考與標籤資料")
    file_quota = st.file_uploader("1️⃣ 監考堂數.xlsx", type=['xlsx'], key=f"f1_{st.session_state['uploader_key']}")
    file_list = st.file_uploader("2️⃣ 監考名單.xlsx", type=['xlsx'], key=f"f2_{st.session_state['uploader_key']}")
    file_type = st.file_uploader("3️⃣ 監考類型總數.xlsx", type=['xlsx'], key=f"f3_{st.session_state['uploader_key']}")
    file_pub = st.file_uploader("4️⃣ 監考總表公布版.xlsx (範本)", type=['xlsx'], key=f"f4_{st.session_state['uploader_key']}")
    file_assign = st.file_uploader("5️⃣ 監考一覽表.xlsx (班級分配範本)", type=['xlsx'], key=f"f5_{st.session_state['uploader_key']}")
    st.write("---")
    file_course = st.file_uploader("6️⃣ 配課表.xlsx (多工作表)", type=['xlsx'], key=f"f6_{st.session_state['uploader_key']}")
    file_label = st.file_uploader("7️⃣ 標籤列印.xlsx (試卷袋範本)", type=['xlsx'], key=f"f7_{st.session_state['uploader_key']}")

with col2:
    st.subheader("⚙️ 2. 考試設定與特許名單")
    selected_sheet = None
    if file_quota:
        xls = pd.ExcelFile(file_quota)
        selected_sheet = st.selectbox("👇 選擇考試項目：", xls.sheet_names)
    
    flex_names = []
    if file_list:
        temp_df = pd.read_excel(file_list, header=None).fillna("")
        teacher_list = temp_df.iloc[2:, 1].astype(str).str.strip().tolist()
        teacher_list = [t for t in teacher_list if t != "" and t != "nan"]
        flex_names = st.multiselect("🛡️ 優先時數不大於名單：", options=teacher_list)

    st.write("")
    c_d1, c_d2 = st.columns(2)
    with c_d1: d1_date = st.date_input("📅 第一天日期：", datetime.now())
    with c_d2: d2_date = st.date_input("📅 第二天日期：", datetime.now())
    
    force_run = st.checkbox("⚠️ 忽略健檢警告，強制執行")
    if st.button("🗑️ 清除所有設定", use_container_width=True):
        st.session_state['results'] = None
        st.session_state['uploader_key'] += 1
        st.rerun()

# ==========================================
# 4. 核心演算法執行
# ==========================================
st.divider()

if st.button("🚀 啟動終極全自動排班系統", type="primary", use_container_width=True):
    if not all([file_quota, file_list, file_type, file_assign]):
        st.error("🚨 請至少確認【1, 2, 3, 5】號基礎檔案皆已上傳！")
    else:
        try:
            # 1. 讀取與計算監考總表
            df_quota = pd.read_excel(file_quota, sheet_name=selected_sheet).fillna("")
            quota_dict = dict(zip(df_quota.iloc[:, 0].astype(str).str.strip(), pd.to_numeric(df_quota.iloc[:, 1], errors='coerce').fillna(0)))
            df_type = pd.read_excel(file_type, header=None).fillna("")
            req_matrix = {'△': [0]*10, '※': [0]*10}
            for i in range(2, len(df_type)):
                row_name = str(df_type.iloc[i, 0]).strip()
                if row_name in ['△', '※']:
                    req_matrix[row_name] = pd.to_numeric(df_type.iloc[i, 1:11], errors='coerce').fillna(0).astype(int).tolist()
            df_list_raw = pd.read_excel(file_list, header=None).fillna("")
            header_df = df_list_raw.iloc[0:2].copy().astype(str).replace('nan', '')
            
            # --- 嘗試自動從【監考一覽表】抓取日期 ---
            df_assign_temp = pd.read_excel(file_assign, header=None).fillna("")
            try:
                sys_d1_str = str(df_assign_temp.iloc[0, 1])[:10].replace('/', '-')
                sys_d2_str = str(df_assign_temp.iloc[0, 6])[:10].replace('/', '-')
            except:
                sys_d1_str = d1_date.strftime('%Y-%m-%d')
                sys_d2_str = d2_date.strftime('%Y-%m-%d')
            
            display_d1 = f"{sys_d1_str[-5:-3]}月{sys_d1_str[-2:]}日"
            display_d2 = f"{sys_d2_str[-5:-3]}月{sys_d2_str[-2:]}日"

            for c in range(3, 8): header_df.iloc[0, c] = display_d1
            for c in range(8, 13): header_df.iloc[0, c] = display_d2
            
            df_list = df_list_raw.iloc[2:].copy()
            teachers = df_list.iloc[:, 1].astype(str).str.strip().tolist()

            with st.spinner("🧠 AI 排班中..."):
                prob = pulp.LpProblem("Scheduling", pulp.LpMinimize)
                vX = {}; vY = {}
                for i in range(len(teachers)):
                    vX[i] = {}; vY[i] = {}
                    for j in range(10):
                        vX[i][j] = pulp.LpVariable(f"X_{i}_{j}", cat='Binary')
                        vY[i][j] = pulp.LpVariable(f"Y_{i}_{j}", cat='Binary')
                penalty = 0
                for i, t in enumerate(teachers):
                    tgt = int(quota_dict.get(t, 0))
                    act = pulp.lpSum([vX[i][k] + vY[i][k]*2 for k in range(10)])
                    prob += act <= tgt
                    dfct = pulp.LpVariable(f"dfct_{i}", 0)
                    prob += act + dfct == tgt
                    penalty += dfct * (1 if t in flex_names else 1000)
                    for j in range(10):
                        prob += vX[i][j] + vY[i][j] <= 1
                        cell_val = str(df_list.iloc[i, j+3]).strip()
                        if cell_val != "" and cell_val != "nan":
                            prob += vX[i][j] == 0; prob += vY[i][j] == 0
                    prob += vX[i][1] >= vY[i][0]
                    prob += vX[i][6] >= vY[i][5]
                for j in range(10):
                    prob += pulp.lpSum([vX[i][j] for i in range(len(teachers))]) == req_matrix['△'][j]
                    prob += pulp.lpSum([vY[i][j] for i in range(len(teachers))]) == req_matrix['※'][j]
                prob += penalty
                prob.solve()

                schedule_dict = {}
                df_out_master = df_list.copy()
                for i, t in enumerate(teachers):
                    res = []
                    for j in range(10):
                        val = str(df_list.iloc[i, j+3]).strip()
                        if val == "" or val == "nan":
                            if vX[i][j].varValue == 1: val = "△"
                            elif vY[i][j].varValue == 1: val = "※"
                            else: val = "" 
                        res.append(val)
                        df_out_master.iloc[i, j+3] = val
                    schedule_dict[t] = res

            # 2. 班級分配邏輯 (計算出分配矩陣)
            class_names_raw = df_assign_temp.iloc[2:, 0].tolist()
            norm_class_names = [clean_str(c) for c in class_names_raw]
            assign_map = {name: idx for idx, name in enumerate(norm_class_names)}
            assigned_matrix = np.empty((len(class_names_raw), 10), dtype=object)

            for day_start in [0, 5]:
                j1 = day_start
                proctors_j1 = [t for t in teachers if schedule_dict[t][j1] in ["△", "※"]]
                random.shuffle(proctors_j1)
                for idx, p in zip(range(len(class_names_raw)), proctors_j1): assigned_matrix[idx, j1] = p
                
                j2 = day_start + 1
                proctors_j2 = [t for t in teachers if schedule_dict[t][j2] in ["△", "※"]]
                bound = {}
                for idx in range(len(class_names_raw)):
                    p_prev = assigned_matrix[idx, j1]
                    if p_prev is not None and schedule_dict[p_prev][j1] == "※" and schedule_dict[p_prev][j2] == "△":
                        assigned_matrix[idx, j2] = p_prev
                        bound[p_prev] = True
                rem = [p for p in proctors_j2 if p not in bound]
                random.shuffle(rem)
                r_idx = 0
                for idx in range(len(class_names_raw)):
                    if assigned_matrix[idx, j2] is None and r_idx < len(rem):
                        assigned_matrix[idx, j2] = rem[r_idx]; r_idx += 1
                for offset in [2, 3, 4]:
                    curr_j = day_start + offset
                    proctors = [t for t in teachers if schedule_dict[t][curr_j] in ["△", "※"]]
                    random.shuffle(proctors)
                    for idx, p in zip(range(len(class_names_raw)), proctors): assigned_matrix[idx, curr_j] = p

            # 3. 生成【監考一覽表】(使用 openpyxl 保留格式)
            with st.spinner("🖨️ 正在將班級分配套入一覽表範本..."):
                wb_assign = openpyxl.load_workbook(file_assign)
                ws_assign = wb_assign.active
                # 填入日期
                if ws_assign.max_row >= 1:
                    ws_assign.cell(row=1, column=2).value = display_d1
                    ws_assign.cell(row=1, column=7).value = display_d2
                # 填入老師名字 (從第 3 列開始)
                for r_idx in range(len(class_names_raw)):
                    for c_idx in range(10):
                        p_name = assigned_matrix[r_idx, c_idx]
                        if p_name:
                            ws_assign.cell(row=r_idx + 3, column=c_idx + 2).value = p_name
                out_assign = io.BytesIO()
                wb_assign.save(out_assign)
                assign_bytes = out_assign.getvalue()

            # 4. 生成【公布版總表】(保留格式)
            pub_bytes = None
            if file_pub:
                with st.spinner("🖨️ 套印公布版總表中..."):
                    wb_pub = openpyxl.load_workbook(file_pub)
                    ws_pub = wb_pub.active
                    h_row = -1; t_cols = []
                    for r in range(1, 20):
                        for c in range(1, ws_pub.max_column + 1):
                            if "教師" in str(ws_pub.cell(row=r, column=c).value): h_row = r; t_cols.append(c)
                        if h_row != -1: break
                    if h_row != -1:
                        for c in t_cols:
                            if h_row - 1 >= 1:
                                ws_pub.cell(row=h_row-1, column=c+2).value = display_d1
                                ws_pub.cell(row=h_row-1, column=c+7).value = display_d2
                            for r in range(h_row+1, ws_pub.max_row + 1):
                                t_v = ws_pub.cell(row=r, column=c).value
                                if t_v and str(t_v).strip() in schedule_dict:
                                    name = str(t_v).strip()
                                    for j in range(5): ws_pub.cell(row=r, column=c+2+j).value = schedule_dict[name][j]
                                    for j in range(5): ws_pub.cell(row=r, column=c+7+j).value = schedule_dict[name][j+5]
                    out_pub = io.BytesIO()
                    wb_pub.save(out_pub)
                    pub_bytes = out_pub.getvalue()

            # 5. 生成【試卷袋標籤】(保留格式)
            label_bytes = None
            if file_course and file_label:
                with st.spinner("🏷️ 合成試卷袋標籤中..."):
                    course_dict = {}
                    xls_c = pd.ExcelFile(file_course)
                    for s_n in xls_c.sheet_names:
                        df_c = pd.read_excel(file_course, sheet_name=s_n, header=None).fillna("")
                        h_idx = 0
                        for r in range(min(5, len(df_c))):
                            row_s = "".join(str(x) for x in df_c.iloc[r, :])
                            if "科目" in row_s or "一1" in row_s: h_idx = r; break
                        clss = [clean_str(x) for x in df_c.iloc[h_idx, :]]
                        for r in range(h_idx + 1, len(df_c)):
                            subj = normalize_subject(df_c.iloc[r, 0])
                            for c_i in range(1, len(clss)):
                                if df_c.iloc[r, c_i]: course_dict[(clss[c_i], subj)] = clean_str(df_c.iloc[r, c_i])
                    
                    wb_l = openpyxl.load_workbook(file_label)
                    ws_l = wb_l.active
                    for r in range(2, ws_l.max_row + 1):
                        vA, vB, vD, vE = ws_l.cell(row=r, column=1).value, ws_l.cell(row=r, column=2).value, ws_l.cell(row=r, column=4).value, ws_l.cell(row=r, column=5).value
                        if not vD: continue
                        cls, subj = clean_str(vD), normalize_subject(vE)
                        teacher = get_teacher_fuzzy(cls, subj, course_dict)
                        if teacher: ws_l.cell(row=r, column=6).value = teacher
                        try: p_val = int(float(str(vA)))
                        except: p_val = -1
                        if cls in assign_map and 1 <= p_val <= 5:
                            d_s = vB.strftime('%Y-%m-%d') if isinstance(vB, datetime) else str(vB).strip().replace('/', '-')
                            off = 0 if sys_d1_str[-5:] in d_s else 5 if sys_d2_str[-5:] in d_s else -1
                            if off != -1:
                                prct = assigned_matrix[assign_map[cls], off + p_val - 1]
                                if prct: ws_l.cell(row=r, column=8).value = prct
                    out_l = io.BytesIO()
                    wb_l.save(out_l)
                    label_bytes = out_l.getvalue()

            st.balloons()
            st.session_state['results'] = {'orig': to_excel_bytes(df_out_master, header_df), 'assign': assign_bytes, 'pub': pub_bytes, 'label': label_bytes}

        except Exception as e:
            st.error(f"發生錯誤: {e}"); st.code(traceback.format_exc())

# ==========================================
# 5. 下載區
# ==========================================
if st.session_state['results']:
    st.divider()
    res = st.session_state['results']
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.download_button("📥 1. 監考總表", res['orig'], "監考總表.xlsx", "application/vnd.ms-excel", use_container_width=True)
    with c2: st.download_button("📥 2. 監考一覽表(保留格式版)", res['assign'], "監考一覽表_分配完成.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary")
    with c3: 
        if res['pub']: st.download_button("📥 3. 公布版套印總表", res['pub'], "公布版總表.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    with c4:
        if res.get('label'): st.download_button("📥 4. 標籤列印(完整版)", res['label'], "標籤列印_完整版.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary")
