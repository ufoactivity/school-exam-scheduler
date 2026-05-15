import streamlit as st
import pandas as pd
import numpy as np
import io
import pulp
import traceback
import random
import openpyxl
import re
from datetime import datetime

# ==========================================
# 1. 網頁頁面配置
# ==========================================
st.set_page_config(page_title="段考監考終極自動化", page_icon="🏫", layout="wide")
st.title("🏫 試務組-段考監考全自動化系統 (終極完全體)")
st.info("💡 終極修復：已解除「監考類型總數」的雷達綁定，解決 index out of bounds 錯誤！現在系統能完美區分靜態數據與動態報表。")

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

def get_teacher_fuzzy(cls, subj, course_dict):
    if (cls, subj) in course_dict: return course_dict[(cls, subj)]
    clean_target = subj.replace('選修', '').replace('彈性學習', '').replace('補強', '').replace('-', '')
    for (c, s), t in course_dict.items():
        if c == cls:
            s_clean = s.replace('選修', '').replace('彈性學習', '').replace('補強', '').replace('-', '')
            if clean_target and (clean_target in s_clean or s_clean in clean_target):
                return t
    return ""

def extract_mm_dd(text, default_month="05"):
    if pd.isna(text) or text is None: return ""
    s = str(text).strip().replace(' ', '')
    m = re.search(r'(\d{1,2})[-/月](\d{1,2})', s)
    if m:
        return f"{int(m.group(1)):02d}-{int(m.group(2)):02d}"
    m2 = re.search(r'(\d{1,2})日', s)
    if m2:
        return f"{int(default_month):02d}-{int(m2.group(1)):02d}"
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
    file_course = st.file_uploader("6️⃣ 配課表.xlsx", type=['xlsx'], key=f"f6_{st.session_state['uploader_key']}")
    file_label = st.file_uploader("7️⃣ 標籤列印.xlsx", type=['xlsx'], key=f"f7_{st.session_state['uploader_key']}")

with col2:
    st.subheader("⚙️ 2. 考試設定與特許名單")
    selected_sheet = None
    if file_quota:
        xls = pd.ExcelFile(file_quota)
        selected_sheet = st.selectbox("👇 選擇考試項目：", xls.sheet_names)
    
    flex_names = []
    if file_list:
        temp_df = pd.read_excel(file_list, header=None).fillna("")
        teacher_list = [str(t).strip() for t in temp_df.iloc[2:, 1] if pd.notna(t) and str(t).strip() != "nan"]
        flex_names = st.multiselect("🛡️ 優先時數不大於名單：", options=teacher_list)

    st.write("")
    c_d1, c_d2 = st.columns(2)
    with c_d1: d1_date = st.date_input("📅 預備日期1：", datetime.now())
    with c_d2: d2_date = st.date_input("📅 預備日期2：", datetime.now())
    
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
        st.error("🚨 請確認必要排考基礎檔案【1, 2, 3, 5】皆已上傳！")
    else:
        try:
            # --- 1. 動態偵測 名單與一覽表 的 AI 節次欄位 ---
            df_list_raw = pd.read_excel(file_list, header=None).fillna("")
            df_assign_temp = pd.read_excel(file_assign, header=None).fillna("")

            def find_ai_10_cols(df):
                for r in range(min(6, len(df))):
                    row_vals = [str(x).strip().split('.')[0] for x in df.iloc[r, :].tolist()]
                    if '1' in row_vals and '2' in row_vals and '3' in row_vals:
                        day1, day2 = [], []
                        cursor = 1
                        for c_idx, val in enumerate(row_vals):
                            if val in ['1', '2', '3', '4', '5']:
                                if val == '1' and len(day1) >= 5: cursor = 2
                                if cursor == 1: day1.append(c_idx)
                                else: day2.append(c_idx)
                        return day1[:5] + day2[:5], r
                return list(range(3, 13)), 2

            list_cols, list_header_r = find_ai_10_cols(df_list_raw)
            assign_cols, assign_periods_r = find_ai_10_cols(df_assign_temp)

            # --- 2. 讀取並建立數學模型限制 ---
            df_quota = pd.read_excel(file_quota, sheet_name=selected_sheet).fillna("")
            quota_dict = dict(zip(df_quota.iloc[:, 0].astype(str).str.strip(), pd.to_numeric(df_quota.iloc[:, 1], errors='coerce').fillna(0)))

            # 【重點修復】：恢復 3 號檔案「監考類型總數」的穩定讀取法
            df_type = pd.read_excel(file_type, header=None).fillna("")
            req_matrix = {'△': [0]*10, '※': [0]*10}
            for i in range(len(df_type)):
                r_name = str(df_type.iloc[i, 0]).strip()
                if r_name in ['△', '※']:
                    req_matrix[r_name] = pd.to_numeric(df_type.iloc[i, 1:11], errors='coerce').fillna(0).astype(int).tolist()

            # --- 3. 精準識別一覽表班級列 ---
            assign_start_idx = -1
            class_keywords = ['商', '國', '電', '資', '廣', '美', '應', '觀', '高', '普', '體']
            for r in range(len(df_assign_temp)):
                v = str(df_assign_temp.iloc[r, 0]).strip()
                if len(v) <= 8 and any(v.startswith(k) for k in class_keywords) and any(c in v for c in ['一', '二', '三', 'ㄧ']):
                    assign_start_idx = r
                    break
            
            class_names_raw = [str(df_assign_temp.iloc[r, 0]).strip() for r in range(assign_start_idx, len(df_assign_temp)) if str(df_assign_temp.iloc[r, 0]).strip() and "註" not in str(df_assign_temp.iloc[r, 0])]
            norm_class_names = [clean_str(c) for c in class_names_raw]
            assign_map = {name: idx for idx, name in enumerate(norm_class_names)}

            # --- 4. PuLP AI 排班運算 ---
            teachers = [str(x).strip() for x in df_list_raw.iloc[list_header_r+1:, 1] if str(x).strip()]
            
            prob = pulp.LpProblem("Exam_Scheduling", pulp.LpMinimize)
            vX = pulp.LpVariable.dicts("X", (range(len(teachers)), range(10)), cat='Binary')
            vY = pulp.LpVariable.dicts("Y", (range(len(teachers)), range(10)), cat='Binary')
            
            penalty = 0
            for i, t in enumerate(teachers):
                tgt = int(quota_dict.get(t, 0))
                prob += pulp.lpSum([vX[i][j] + vY[i][j]*2 for j in range(10)]) <= tgt
                dfct = pulp.LpVariable(f"dfct_{i}", 0)
                prob += pulp.lpSum([vX[i][j] + vY[i][j]*2 for j in range(10)]) + dfct == tgt
                penalty += dfct * (1 if t in flex_names else 1000)
                for j in range(10):
                    prob += vX[i][j] + vY[i][j] <= 1
                    cv = str(df_list_raw.iloc[i+list_header_r+1, list_cols[j]]).strip()
                    if cv and cv != "nan": prob += vX[i][j] == 0; prob += vY[i][j] == 0
                prob += vX[i][1] >= vY[i][0]
                prob += vX[i][6] >= vY[i][5]
            for j in range(10):
                prob += pulp.lpSum([vX[i][j] for i in range(len(teachers))]) == req_matrix['△'][j]
                prob += pulp.lpSum([vY[i][j] for i in range(len(teachers))]) == req_matrix['※'][j]
            prob += penalty
            prob.solve()

            # 寫入有 AI 符號的監考總表
            df_out_master = df_list_raw.copy()
            schedule_dict = {}
            for i, t in enumerate(teachers):
                res = []
                for j in range(10):
                    v = str(df_list_raw.iloc[i+list_header_r+1, list_cols[j]]).strip()
                    if v == "" or v == "nan":
                        if vX[i][j].varValue == 1: v = "△"
                        elif vY[i][j].varValue == 1: v = "※"
                        else: v = ""
                    res.append(v)
                    df_out_master.iloc[i+list_header_r+1, list_cols[j]] = v
                schedule_dict[t] = res

            # --- 5. 班級全自動分發與寫入一覽表 ---
            assigned_matrix = np.empty((len(class_names_raw), 10), dtype=object)
            for day_start in [0, 5]:
                j1 = day_start
                avail_j1 = [t for t in teachers if schedule_dict[t][j1] in ["△", "※"]]
                random.shuffle(avail_j1)
                for idx, p in zip(range(len(class_names_raw)), avail_j1): assigned_matrix[idx, j1] = p
                
                j2 = day_start + 1
                avail_j2 = [t for t in teachers if schedule_dict[t][j2] in ["△", "※"]]
                bound = {}
                for idx in range(len(class_names_raw)):
                    prev = assigned_matrix[idx, j1]
                    if prev and schedule_dict[prev][j1] == "※" and schedule_dict[prev][j2] == "△":
                        assigned_matrix[idx, j2] = prev; bound[prev] = True
                rem = [p for p in avail_j2 if p not in bound]
                random.shuffle(rem)
                r_ptr = 0
                for idx in range(len(class_names_raw)):
                    if assigned_matrix[idx, j2] is None and r_ptr < len(rem):
                        assigned_matrix[idx, j2] = rem[r_ptr]; r_ptr += 1
                for offset in [2, 3, 4]:
                    curr_j = day_start + offset
                    avail = [t for t in teachers if schedule_dict[t][curr_j] in ["△", "※"]]
                    random.shuffle(avail)
                    for idx, p in zip(range(len(class_names_raw)), avail): assigned_matrix[idx, curr_j] = p

            wb_assign = openpyxl.load_workbook(file_assign)
            ws_assign = wb_assign.active
            for r_idx in range(len(class_names_raw)):
                target_r = assign_start_idx + r_idx + 1
                for c_idx in range(10):
                    p_name = assigned_matrix[r_idx, c_idx]
                    if p_name:
                        target_c = assign_cols[c_idx] + 1
                        cell = ws_assign.cell(row=target_r, column=target_c)
                        if type(cell).__name__ != 'MergedCell':
                            cell.value = str(p_name)
            
            out_assign = io.BytesIO()
            wb_assign.save(out_assign)
            assign_bytes = out_assign.getvalue()

            # --- 6. 公布版套印 (自動對位) ---
            pub_bytes = None
            if file_pub:
                with st.spinner("🖨️ 正在無縫套印公布版..."):
                    wb_pub = openpyxl.load_workbook(file_pub)
                    ws_pub = wb_pub.active
                    
                    pub_periods_row = -1
                    for r in range(1, min(20, ws_pub.max_row + 1)):
                        r_vals = [str(ws_pub.cell(row=r, column=c).value).strip().split('.')[0] for c in range(1, ws_pub.max_column + 1)]
                        if '1' in r_vals and '2' in r_vals and '3' in r_vals:
                            pub_periods_row = r; break
                    
                    if pub_periods_row != -1:
                        pub_target_cols = []
                        cursor = 1
                        d1_c, d2_c = [], []
                        for c in range(1, ws_pub.max_column + 1):
                            val = str(ws_pub.cell(row=pub_periods_row, column=c).value).strip().split('.')[0]
                            if val in ['1', '2', '3', '4', '5']:
                                if val == '1' and len(d1_c) >= 5: cursor = 2
                                if cursor == 1: d1_c.append(c)
                                else: d2_c.append(c)
                        pub_target_cols = d1_c[:5] + d2_c[:5]
                        
                        h_row = -1; t_cols = []
                        for r in range(1, min(20, ws_pub.max_row + 1)):
                            for c in range(1, ws_pub.max_column + 1):
                                if "教師" in str(ws_pub.cell(row=r, column=c).value): h_row = r; t_cols.append(c)
                            if h_row != -1: break
                        
                        if h_row != -1 and len(pub_target_cols) >= 10:
                            for c in t_cols:
                                for r in range(h_row+1, ws_pub.max_row + 1):
                                    t_val = ws_pub.cell(row=r, column=c).value
                                    if t_val and str(t_val).strip() in schedule_dict:
                                        name = str(t_val).strip()
                                        for j in range(10):
                                            cw = ws_pub.cell(row=r, column=pub_target_cols[j])
                                            if type(cw).__name__ != 'MergedCell': cw.value = schedule_dict[name][j]
                    out_pub = io.BytesIO()
                    wb_pub.save(out_pub)
                    pub_bytes = out_pub.getvalue()

            # --- 7. 試卷袋標籤合成 (智慧日期匹配 + 手排100%複製) ---
            label_bytes = None
            if file_course and file_label:
                with st.spinner("🏷️ 啟動【所見即所得】標籤精準合成..."):
                    course_dict = {}
                    xls_course = pd.ExcelFile(file_course)
                    for sheet in xls_course.sheet_names:
                        df_c = pd.read_excel(file_course, sheet_name=sheet, header=None).fillna("")
                        h_idx = 0
                        for r in range(min(5, len(df_c))):
                            row_str = "".join(str(x) for x in df_c.iloc[r, :])
                            if "科目" in row_str or "一1" in row_str: h_idx = r; break
                        clss = [clean_str(x) for x in df_c.iloc[h_idx, :]]
                        for r in range(h_idx + 1, len(df_c)):
                            subj = normalize_subject(df_c.iloc[r, 0])
                            for c_i in range(1, len(clss)):
                                if df_c.iloc[r, c_i]: course_dict[(clss[c_i], subj)] = clean_str(df_c.iloc[r, c_i])
                    
                    date_row_idx = max(0, assign_periods_r - 1)
                    detected_month = "05"
                    for c in range(1, ws_assign.max_column + 1):
                        d_v = str(ws_assign.cell(row=date_row_idx + 1, column=c).value)
                        m_idx = re.search(r'(\d{1,2})[-/月]', d_v)
                        if m_idx: detected_month = m_idx.group(1); break

                    schedule_col_map = {}
                    curr_d = ""
                    for c in range(1, ws_assign.max_column + 1):
                        d_val = ws_assign.cell(row=date_row_idx + 1, column=c).value
                        p_val = ws_assign.cell(row=assign_periods_r + 1, column=c).value
                        if d_val is not None and str(d_val).strip() not in ["", "None"]:
                            m = extract_mm_dd(d_val, default_month=detected_month)
                            if m: curr_d = m
                        if curr_d and p_val is not None and str(p_val).strip() not in ["", "None"]:
                            p_str = str(p_val).strip().split('.')[0]
                            schedule_col_map[(curr_d, p_str)] = c

                    wb_label = openpyxl.load_workbook(file_label)
                    ws_label = wb_label.active
                    col_map = {}
                    for c in range(1, ws_label.max_column + 1):
                        val = clean_str(ws_label.cell(row=1, column=c).value)
                        if val: col_map[val] = c
                    col_teacher = col_map.get('任課教師', 6)  
                    col_proctor = col_map.get('監考老師', 8)  

                    for r in range(2, ws_label.max_row + 1):
                        val_A = str(ws_label.cell(row=r, column=1).value or "").strip()
                        val_B = ws_label.cell(row=r, column=2).value
                        val_D = str(ws_label.cell(row=r, column=4).value or "").strip()
                        val_E = str(ws_label.cell(row=r, column=5).value or "").strip()
                        
                        if not val_D: continue
                        cls, subj = clean_str(val_D), normalize_subject(val_E)
                        
                        teacher = get_teacher_fuzzy(cls, subj, course_dict)
                        if teacher:
                            t_c = ws_label.cell(row=r, column=col_teacher)
                            if type(t_c).__name__ != 'MergedCell': t_c.value = teacher
                        
                        try: p_val_str = str(int(float(val_A)))
                        except: p_val_str = ""
                        l_date = val_B.strftime('%m-%d') if isinstance(val_B, datetime) else extract_mm_dd(str(val_B), default_month=detected_month)
                        
                        if cls in assign_map and l_date and p_val_str:
                            target_c = schedule_col_map.get((l_date, p_val_str))
                            if target_c:
                                target_r = assign_start_idx + assign_map[cls] + 1
                                proctor = ws_assign.cell(row=target_r, column=target_c).value
                                if proctor and str(proctor).strip() != 'None':
                                    p_c = ws_label.cell(row=r, column=col_proctor)
                                    if type(p_c).__name__ != 'MergedCell': p_c.value = str(proctor)

                    out_label = io.BytesIO()
                    wb_label.save(out_label)
                    label_bytes = out_label.getvalue()

            st.balloons()
            st.session_state['results'] = {
                'orig': to_excel_bytes(df_out_master, None),
                'assign': assign_bytes,
                'pub': pub_bytes,
                'label': label_bytes
            }
            st.success("🎉 旗艦完全體版排班完成！所有功能已完美歸位。")

        except Exception as e:
            st.error(f"發生錯誤: {e}"); st.code(traceback.format_exc())

# ==========================================
# 5. 下載區
# ==========================================
if st.session_state['results']:
    st.divider()
    res = st.session_state['results']
    c1, c2, c3, c4 = st.columns(4)
    with c1: 
        st.download_button("📥 1. 監考總表", res['orig'], "監考總表.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    with c2: 
        st.download_button("📥 2. 監考一覽表(保留格式版)", res['assign'], "監考一覽表_分配完成.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary")
    with c3: 
        if res.get('pub'): 
            st.download_button("📥 3. 公布版套印總表", res['pub'], "公布版總表.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    with c4:
        if res.get('label'): 
            st.download_button("📥 4. 標籤列印(完整版)", res['label'], "標籤列印_完整版.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary")
