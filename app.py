import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from docx import Document
from io import BytesIO

# --- 初始化防呆鎖定狀態 (Session State) ---
if 'setup_complete' not in st.session_state:
    st.session_state.setup_complete = False
if 'df' not in st.session_state:
    st.session_state.df = None
if 'date_str' not in st.session_state:
    st.session_state.date_str = datetime.today().strftime("%Y-%m-%d")

# 車隊人員名單
ALL_MEMBERS = [
    "丁秋吟", "王永慶", "王銓德", "王志文", "王家業", "朱家樺", "江旼珀", "吳上苑", "吳宜汶", "呂恩昕", 
    "呂淑惠", "李國誥", "李榮斌", "李穎裕", "阮智偉", "周昆佑", "周志暐", "周志祥", "林志嶸", "林永松", 
    "林志佳", "劉柔君", "林佳宏", "林國芳", "林婉茹", "林智偉", "林華盛", "林瑞華", "林瑞燿", "林嘉信", 
    "邱保銘", "邱信培", "徐琮凱", "柯富強", "洪偉倫", "胡中興", "胡耀仁", "夏進通", "涂旻聖", "張進源", 
    "張文男", "張世明", "張仕欣", "張正中", "張百江", "張孟哲", "張錦升", "張勝富", "張路雄", "張鈞銘", 
    "張仕宗", "張聰捷", "梁善鈞", "郭政富", "陳文政", "陳仕明", "陳怡如", "陳瑜玲", "陳冠良", "陳春文", 
    "陳敏訓", "陳盛宏", "陳進忠", "陳佳忠", "曾雅惠", "游正豪", "游育民", "游振和", "黃世政", "黃明興", 
    "黃信隆", "黃冠霖", "黃堃珉", "黃賀進", "黃麗萍", "楊俊逸", "魏捷祥", "楊閔森", "廖崇凱", "廖竣傑", 
    "廖致維", "熊致堯", "蔡芷纭", "蔡玉雯", "蔡榮峰", "鄭銘宗", "蕭大勛", "蕭鈺潔", "謝仁政", "魏志龍", 
    "藍慧真", "詹獻睿", "詹孟杰", "曾舜麟", "謝明佑", "羅靜宜", "姜小平", "莊麗雪", "吳慧芳", "蔡瑞賓", 
    "張上觀", "陳素貞", "吳建興", "呂昇印", "孟繁光", "林坤茂", "邱榮家", "徐偉欽", "鄭淵太", "張志州", 
    "張富山", "陳珀升", "陳溪宗", "陳瑋楊", "曾明雄", "詹憲國", "廖宏輝", "廖翊均", "劉邦杰", "劉明煌", 
    "蔡榮祥", "蔡榮華", "周智勤", "謝志忠", "涂欽耀", "張志仲", "蕭森巍", "張銀恭", "王正錄", "曾建勳", 
    "黃智煒", "劉海森", "賴南君", "劉權漢", "游伊君", "陳勇志", "陳文田", "林秋雄", "賴永昌", "劉煉騰", 
    "林錦志", "林明忠", "張哲誠", "詹昆學", "陳彥宏", "許馨云", "張橋語"
]

def generate_word_report(date_str, attendees):
    doc = Document()
    doc.add_heading('木工機械單車協會 - 團騎點名紀錄', 0)
    doc.add_heading(f'日期：{date_str}', level=1)
    doc.add_paragraph(f'本次參與總人數：{len(attendees)} 人')
    doc.add_heading('出席名單：', level=2)
    doc.add_paragraph("、".join(attendees))
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

st.title("🚴‍♂️ 車隊團騎點名系統")

# ==========================================
# 階段一：設定畫面 (完成後會自動隱藏)
# ==========================================
if not st.session_state.setup_complete:
    st.info("💡 請先完成下方設定。確認後介面會自動鎖定並隱藏，避免點名時誤觸。")
    
    st.markdown("### 步驟一：載入資料")
    mode = st.radio("請選擇資料載入方式：", ["上傳舊有總表接續點名", "建立全新總表"])
    
    temp_df = None
    if mode == "上傳舊有總表接續點名":
        uploaded_file = st.file_uploader("📂 選擇先前的 Excel 總表", type=["xlsx"])
        if uploaded_file is not None:
            temp_df = pd.read_excel(uploaded_file)
            st.success("✅ 舊表單載入成功！")
    else:
        temp_df = pd.DataFrame({"姓名": ALL_MEMBERS})
        st.info("🆕 將建立全新表單。")

    st.markdown("### 步驟二：確認點名日期")
    date_options = [(datetime.today() + timedelta(days=i)).strftime("%Y-%m-%d") for i in range(-14, 15)]
    selected_date = st.selectbox("📅 選擇團騎日期：", date_options, index=14)

    # 鎖定按鈕
    if st.button("🔒 確認設定並開始點名"):
        if mode == "上傳舊有總表接續點名" and temp_df is None:
            st.warning("⚠️ 請先上傳檔案再繼續！")
        else:
            st.session_state.df = temp_df
            st.session_state.date_str = selected_date
            st.session_state.setup_complete = True
            st.rerun() # 重新整理畫面，進入點名模式

# ==========================================
# 階段二：點名主畫面 (設定完成後才會顯示)
# ==========================================
else:
    st.success(f"📌 目前鎖定點名日期：**{st.session_state.date_str}**")
    
    st.markdown("### 步驟三：人員點名")
    st.write("🎙️ **語音/注音快搜**：點擊下方框框，直接用手機鍵盤打注音，或按鍵盤上的「麥克風🎤」直接唸名字！")
    
    attendees = st.multiselect("✅ 請勾選今日出席人員：", ALL_MEMBERS)
    
    if st.button("💾 完成點名並產出報表"):
        df = st.session_state.df
        date_str = st.session_state.date_str
        
        if not attendees:
            st.warning("⚠️ 請至少選擇一位出席人員！")
        elif date_str in df.columns and not df[df[date_str] == 'V'].empty:
            st.warning(f"⚠️ {date_str} 已經有點名紀錄囉！")
        else:
            # 補齊新名單
            missing_members = [m for m in ALL_MEMBERS if m not in df['姓名'].values]
            if missing_members:
                new_rows = pd.DataFrame({"姓名": missing_members})
                df = pd.concat([df, new_rows], ignore_index=True)

            # 打 V
            if date_str not in df.columns:
                df[date_str] = ""
            for member in attendees:
                df.loc[df['姓名'] == member, date_str] = "V"

            # --- Excel 格式整理 ---
            if '編號' in df.columns:
                df = df.drop(columns=['編號'])
            if '總次數' in df.columns:
                df = df.drop(columns=['總次數'])

            date_cols = [col for col in df.columns if col != '姓名']
            df['總次數'] = df[date_cols].apply(lambda x: (x == 'V').sum(), axis=1)

            # 加入依序編號
            df.insert(0, '編號', range(1, len(df) + 1))
            final_cols = ['編號', '姓名'] + date_cols + ['總次數']
            df = df[final_cols]
            
            # 更新 Session 內的資料
            st.session_state.df = df

            st.success(f"🎉 已成功記錄 {len(attendees)} 人！請務必下載下方檔案。")

            excel_buffer = BytesIO()
            df.to_excel(excel_buffer, index=False)
            word_bytes = generate_word_report(date_str, attendees)

            col1, col2 = st.columns(2)
            with col1:
                st.download_button("📥 下載 Excel 總表", data=excel_buffer.getvalue(), file_name=f"車隊點名總表_{date_str}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with col2:
                st.download_button("📥 下載 Word 紀錄檔", data=word_bytes, file_name=f"{date_str}_團騎點名紀錄.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            
            st.write("📊 總表預覽：")
            st.dataframe(df)

    # 提供解除鎖定的後門
    st.markdown("---")
    if st.button("⚙️ 修改日期或重新載入表單"):
        st.session_state.setup_complete = False
        st.rerun()
