import streamlit as st
import pandas as pd
import os
import altair as alt
import re
from pptx import Presentation
from datetime import datetime
from dateutil.relativedelta import relativedelta
import io

# --- 準備：アプリで使うフォルダとファイルの場所を設定 ---
os.makedirs("exports", exist_ok=True)
MASTER_FILE = "master_data.xlsx"
TEMPLATE_PATH = "template_quote.pptx"

# アプリの基本設定（タイトルや画面幅）
st.set_page_config(page_title="Scope 3見積シミュレーション", layout="wide")

# --- 準備：Excelからマスタデータを読み込む機能 ---
def load_excel_data():
    try:
        df = pd.read_excel(MASTER_FILE, sheet_name="ServiceMaster")
        multi_df = pd.read_excel(MASTER_FILE, sheet_name="GroupMultipliers")
        scale_df = pd.read_excel(MASTER_FILE, sheet_name="ScaleMultipliers")
        
        df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
        if len(df.columns) >= 5:
            df.rename(columns={df.columns[4]: 'Description'}, inplace=True)
            
        return df, multi_df, scale_df
    except Exception as e:
        st.error(f"Excelの読み込みに失敗しました: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# 読み込みの実行
df_master, df_multi, df_scale = load_excel_data()

# --- デザイン：画面の見た目を整える設定 (CSS) ---
st.markdown("""
    <style>
    /* 全体の背景色 */
    .stApp { background-color: #fcfaf5; }
    
    /* セグメントごとの見出しデザイン */
    .section-header { 
        padding: 15px; border-radius: 10px; color: white; margin-top: 30px; margin-bottom: 15px; 
        font-weight: bold; font-size: 1.3em; box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .common-header { background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); border-bottom: 4px solid #162a50; }
    .upstream-header { background: linear-gradient(135deg, #b21f1f 0%, #f44336 100%); border-bottom: 4px solid #7f1616; }
    .downstream-header { background: linear-gradient(135deg, #fbc02d 0%, #fdfc47 100%); border-bottom: 4px solid #c49000; color: #333 !important; }
    
    /* アコーディオンのデザイン */
    .stExpander { border: 1px solid #e0e0e0; background-color: white; margin-bottom: 5px; border-radius: 8px; }
    
    /* 説明文ボックスのデザイン */
    .desc-box {
        background-color: #eef2f7 !important;
        border-left: 5px solid #2a5298 !important;
        padding: 10px 15px !important;
        margin: 5px 0 15px 35px !important;
        border-radius: 0 8px 8px 0 !important;
        font-size: 0.88em !important;
        color: #334155 !important;
        line-height: 1.5 !important;
    }

    /* 金額表示コンテナのデザイン */
    .price-container {
        background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
        padding: 25px; border-radius: 15px; color: white; text-align: center; 
        margin-top: 10px; box-shadow: 0 6px 12px rgba(0,0,0,0.15);
    }
    .price-net { font-size: 45px; color: #ffffff; font-weight: bold; text-shadow: 1px 1px 2px rgba(0,0,0,0.2); }
    .price-tax { font-size: 1.2em; color: #f0f0f0; margin-top: 5px; font-weight: 500; }

    /* CSVボタンのカスタムデザイン */
    div.stDownloadButton > button {
        background-color: #1e3c72 !important;
        color: white !important;
        border-radius: 12px !important;
        padding: 15px 30px !important;
        font-size: 1.2em !important;
        font-weight: bold !important;
        border: 2px solid #162a50 !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1) !important;
        transition: all 0.3s ease !important;
    }
    div.stDownloadButton > button:hover {
        background-color: #2a5298 !important;
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 15px rgba(0,0,0,0.2) !important;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 内部関数：一括選択ボタンの挙動 ---
def toggle_group_all(group_name, key):
    new_state = st.session_state[key]
    g_df = df_master[df_master["Group"] == group_name]
    for _, row in g_df.iterrows():
        st.session_state[f"task_{row['Category']}_{row['Task']}"] = new_state
        st.session_state[f"all_cat_{row['Category']}"] = new_state

def toggle_category_all(cat_name, key):
    new_state = st.session_state[key]
    c_df = df_master[df_master["Category"] == cat_name]
    for _, row in c_df.iterrows():
        st.session_state[f"task_{row['Category']}_{row['Task']}"] = new_state

# --- サイドバー：基本情報の入力エリア ---
with st.sidebar:
    st.header("⚙️ 基本設定")
    company_name = st.text_input("会社名", value="〇〇株式会社")
    start_date = st.date_input("支援開始予定月", datetime.now())
    hourly_rate = st.number_input("時間単価 (円)", value=40000, step=1000)

    if not df_scale.empty:
        scale_options = dict(zip(df_scale['ScaleName'], df_scale['Multiplier']))
    else:
        scale_options = {"中小企業": 1.0}
    
    company_scale = st.selectbox("企業規模", list(scale_options.keys()), index=0)
    multiplier = scale_options[company_scale]
    
    st.divider()
    company_count = st.select_slider("グループ会社数", options=[0, 1, 2, 3, 4, 5, 6],
                                     format_func=lambda x: f"{x}社" if x <= 5 else "5社超")
    
    if not df_multi.empty:
        multi_row = df_multi[df_multi['CompanyCount'] == company_count].iloc[0]
        group_multiplier = multi_row['Multiplier']
        is_special_case = company_count > 5
    else:
        group_multiplier = 1.0
        is_special_case = False
        
    st.divider()
    region_type = st.radio("対象地域", ["国内のみ", "海外含む"])
    is_eng = st.checkbox("成果物の英語提出あり (+10h)") if region_type == "海外含む" else False
    english_hours = 10 if is_eng else 0
            
    st.divider()
    duration_months = st.slider("支援期間 (ヶ月)", 1, 12, 6)
    end_date = start_date + relativedelta(months=duration_months)
    
    mtg_freq = st.number_input("定期MTG回数 / 月", value=2)
    workshop_count = st.number_input("勉強会開催回数", value=1, max_value=2 if company_count > 0 else 5)

    fixed_hours = (duration_months * mtg_freq * 1.0) + (workshop_count * 5.0) + english_hours

# --- メイン画面：タスク選択エリア ---
st.title("🌱 Scope 3算定支援コンサルティング見積シミュレーション")

total_base_hours = fixed_hours 
selected_tasks_list = []

# 固定項目の自動集計
selected_tasks_list.append({"Category": "その他", "Task": "キックオフ", "Hours":})
selected_tasks_list.append({"Category": "その他", "Task": "定期MTG", "Hours": duration_months * mtg_freq})
if workshop_count > 0:
    selected_tasks_list.append({"Category": "その他", "Task": "勉強会", "Hours": workshop_count * 5.0})
if english_hours > 0:
    selected_tasks_list.append({"Category": "その他", "Task": "英語対応", "Hours": 10.0})

# セグメント別タスクの表示
if not df_master.empty:
    for group in ["共通", "上流", "下流"]:
        h_class = "common-header" if group == "共通" else "upstream-header" if group == "上流" else "downstream-header"
        st.markdown(f'<div class="section-header {h_class}">{group}セグメント</div>', unsafe_allow_html=True)
        
        g_key = f"g_all_{group}"
        st.checkbox(f"【{group}】を一括選択", key=g_key, on_change=toggle_group_all, args=(group, g_key))
        
        g_df = df_master[df_master["Group"] == group]
        cols = st.columns(2)
        cat_list = g_df["Category"].unique()
        
        for idx, cat_name in enumerate(cat_list):
            c_df = g_df[g_df["Category"] == cat_name]
            with cols[idx % 2]:
                for _, r in c_df.iterrows():
                    t_key = f"task_{cat_name}_{r['Task']}"
                    if t_key not in st.session_state:
                        st.session_state[t_key] = r['Required']

                selected_count = sum([st.session_state.get(f"task_{cat_name}_{r['Task']}", False) for _, r in c_df.iterrows()])
                
                if selected_count == len(c_df):
                    display_label = f"📁 {cat_name} （✅ 全選択中）"
                elif selected_count > 0:
                    display_label = f"📁 {cat_name} （🔹 {selected_count}/{len(c_df)} 選択中）"
                else:
                    display_label = f"📁 {cat_name} （未選択）"

                is_expanded = selected_count > 0
                with st.expander(display_label, expanded=is_expanded):
                    c_key = f"all_cat_{cat_name}"
                    st.session_state[c_key] = (selected_count == len(c_df))
                    st.checkbox(f"└ {cat_name}を一括選択", key=c_key, on_change=toggle_category_all, args=(cat_name, c_key))
                    st.divider()
                    
                    for _, row in c_df.iterrows():
                        t_key = f"task_{row['Category']}_{row['Task']}"
                        base_h = row["Hours"]
                        calc_h = base_h * group_multiplier if (company_count > 0 and row["Group"] != "共通") else base_h
                        
                        is_checked = st.checkbox(f"　{row['Task']} ({calc_h:.1f}h)", key=t_key)
                        desc_text = str(row.get('Description', '')).strip()
                        if desc_text and desc_text != 'nan' and desc_text != '':
                            st.markdown(f'<div class="desc-box">{desc_text}</div>', unsafe_allow_html=True)

                        if is_checked:
                            total_base_hours += calc_h
                            display_cat = "その他" if (cat_name.startswith("0") or not cat_name.startswith("C")) else cat_name
                            selected_tasks_list.append({
                                "Category": display_cat, 
                                "Task": row['Task'], 
                                "Hours": calc_h,
                                "Description": desc_text if desc_text != 'nan' else ""
                            })

# --- 画面表示：現在の選択タスク一覧 ---
if selected_tasks_list and not is_special_case:
    st.divider()
    summary_df = pd.DataFrame(selected_tasks_list)
    
    def sort_cats(c):
        if c == "その他": return 999
        num = re.findall(r'\d+', c)
        return int(num[0]) if num else 998
    unique_cats = sorted(summary_df['Category'].unique(), key=sort_cats)

    html = '<div style="background-color:#ffffff;border:1px solid #e2e8f0;border-radius:12px;padding:20px;margin-bottom:25px;box-shadow:0 2px 4px rgba(0,0,0,0.05);">'
    html += '<div style="margin-bottom:15px;font-weight:bold;color:#1e3c72;font-size:1.1em;border-bottom:2px solid #f1f5f9;padding-bottom:10px;">📝 現在の選択タスク一覧（合計 ' + str(len(selected_tasks_list)) + ' 項目）</div>'
    for cat in unique_cats:
        tasks = summary_df[summary_df['Category'] == cat]['Task'].tolist()
        tasks_str = " ／ ".join(tasks)
        line = '<div style="display:flex;margin-bottom:12px;border-bottom:1px solid #f8fafc;padding-bottom:8px;">'
        line += '<div style="flex:0 0 150px;font-weight:bold;color:#2a5298;font-size:0.85em;background-color:#f1f5f9;padding:4px 8px;border-radius:6px;text-align:left;align-self:flex-start;">' + str(cat) + '</div>'
        line += '<div style="flex:1;font-size:0.9em;color:#334155;margin-left:15px;line-height:1.6;text-align:left;">' + str(tasks_str) + '</div>'
        line += '</div>'
        html += line
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)

# --- 画面表示：見積金額の計算結果 ---
adj_h = total_base_hours * multiplier
net_price = adj_h * hourly_rate
tax_price = net_price * 1.1

if is_special_case:
    st.markdown('<div style="background-color: #EB5228; color: white; padding: 20px; border-radius: 10px; text-align: center; font-size: 1.5em; font-weight: bold; margin-top: 20px;">個別見積（SAへ要相談）</div>', unsafe_allow_html=True)
else:
    st.markdown(f"""
        <div class="price-container">
            <p style="margin: 0; font-size: 1.0em; opacity: 0.9;">御見積合計金額 (税抜)</p>
            <div class="price-net">¥{int(net_price):,}</div>
            <div class="price-tax">(税込 ¥{int(tax_price):,})</div>
            <p style="margin-top: 15px; font-size: 0.85em; opacity: 0.85;">
                合計工数: {total_base_hours:.1f}h / 調整後工数: {adj_h:.1f}h
            </p>
        </div>
        <div style="margin-bottom: 60px;"></div>
        """, unsafe_allow_html=True)

# --- 画面表示：見積内訳の可視化グラフ ---
st.header("📊 見積内訳分析")
st.markdown("<div style='margin-bottom: 25px;'></div>", unsafe_allow_html=True)

if selected_tasks_list and not is_special_case:
    viz_df = pd.DataFrame(selected_tasks_list)
    viz_df['Price'] = viz_df['Hours'] * multiplier * hourly_rate
    cat_summary = viz_df.groupby('Category')['Price'].sum().reset_index()
    
    def shorten_name(name):
        match = re.search(r'(C\d+)', name)
        return match.group(1) if match else name
    
    cat_summary['DisplayCategory'] = cat_summary['Category'].apply(shorten_name)
    total_val = cat_summary['Price'].sum()
    cat_summary['割合(%)'] = (cat_summary['Price'] / total_val * 100).round(1)

    def get_sort_key(cat_text):
        if cat_text == "その他": return -1
        nums = re.findall(r'\d+', cat_text)
        return int(nums[0]) if nums else 999
    
    cat_summary['sort_val'] = cat_summary['Category'].apply(get_sort_key)
    cat_summary = cat_summary.sort_values('sort_val').reset_index(drop=True)

    col_chart, col_table = st.columns([2, 1])
    
    with col_chart:
        chart = alt.Chart(cat_summary).mark_bar(
            cornerRadiusTopLeft=2, cornerRadiusTopRight=2, size=30
        ).encode(
            x=alt.X('DisplayCategory:N', sort=None, title=None, 
                    axis=alt.Axis(labelAngle=0, labelColor='#1a202c', domainColor='#000000', domainWidth=1.5)),
            y=alt.Y('Price:Q', title='金額 (円)', 
                    axis=alt.Axis(grid=False, domainColor='#000000', domainWidth=1.5, titleAnchor='end')),
            color=alt.Color('Price:Q', scale=alt.Scale(range=['#cbd5e0', '#1a202c']), legend=None),
            tooltip=['Category', 'Price', '割合(%)']
        ).properties(height=350, background='#fcfaf5').configure_view(strokeWidth=0).configure_axis(ticks=False, labelFontSize=11, titleFontSize=12)
        st.altair_chart(chart, use_container_width=True)

    with col_table:
        st.markdown("<div style='margin-top: 5px;'></div>", unsafe_allow_html=True)
        st.write("💰 カテゴリ別内訳")
        formatted_summary = cat_summary.copy()
        formatted_summary['金額'] = formatted_summary['Price'].apply(lambda x: f"¥{int(x):,}")
        formatted_summary['比率'] = formatted_summary['割合(%)'].apply(lambda x: f"{x}%")
        st.dataframe(formatted_summary[['Category', '金額', '比率']], hide_index=True, use_container_width=True)

    # --- CSV出力：ファイル名生成ルールの修正反映 ---
    st.markdown("<br>", unsafe_allow_html=True)
    _, btn_col, _ = st.columns([1, 2, 1])
    with btn_col:
        # 基本情報ヘッダー
        basic_info = [
            ["項目", "設定値"],
            ["会社名", company_name],
            ["支援開始予定月", start_date.strftime('%Y年%m月')],
            ["支援終了予定月", end_date.strftime('%Y年%m月')],
            ["支援期間", f"{duration_months}ヶ月"],
            ["企業規模", company_scale],
            ["企業規模係数", f"x {multiplier}"],
            ["グループ会社数", f"{company_count}社" if company_count <= 5 else "5社超"],
            ["対象地域", region_type],
            ["英語対応", "あり" if is_eng else "なし"],
            ["時間単価", f"¥{hourly_rate:,}"],
            ["合計工数(調整前)", f"{total_base_hours:.1f}h"],
            ["合計工数(調整後)", f"{adj_h:.1f}h"],
            ["合計金額(税抜)", f"¥{int(net_price):,}"],
            ["合計金額(税込)", f"¥{int(tax_price):,}"],
            ["", ""], 
            ["【内訳詳細】", ""],
            ["カテゴリ", "タスク名", "工数(h) ※規模係数適用済", "時間単価(円)", "内訳金額(円)", "内容説明"]
        ]
        
        # 内訳詳細データ
        details = []
        for item in selected_tasks_list:
            adjusted_task_hours = item['Hours'] * multiplier
            task_price = int(adjusted_task_hours * hourly_rate)
            details.append([
                item['Category'], 
                item['Task'], 
                round(adjusted_task_hours, 2),
                int(hourly_rate),
                task_price,
                item.get('Description', '')
            ])
        
        csv_buffer = io.StringIO()
        pd.DataFrame(basic_info).to_csv(csv_buffer, index=False, header=False)
        pd.DataFrame(details).to_csv(csv_buffer, index=False, header=False)
        csv_output = csv_buffer.getvalue().encode('utf_8_sig')
        
        # ファイル名の生成
        today_str = datetime.now().strftime('%Y%m%d')
        file_name_full = f"{today_str}_Scope3見積_{company_name}.csv"
        
        st.download_button(
            label="📥 見積報告書(CSV)を出力する",
            data=csv_output,
            file_name=file_name_full,
            mime="text/csv",
            use_container_width=True,

        )


