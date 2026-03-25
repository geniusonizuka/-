import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
from io import BytesIO

# --- 初始化防呆與狀態 (Session State) ---
if 'setup_complete' not in st.session_state:
    st.session_state.setup_complete = False
if 'df' not in st.session_state:
    st.session_state.df = None
if 'date_str' not in st.session_state:
    st.session_state.date_str = None
if 'location' not in st.session_state:
    st.session_state.location = ""
if 'attendees' not in st.session_state:
    st.session_state.attendees = [] # 儲存已簽到人員
if 'report_generated' not in st.session_state:
    st.session_state.report_generated = False 
if 'excel_data' not in st.session_state:
    st.session_state.excel_data = None
if 'word_data' not in st.session_state:
    st.session_state.word_data = None
if 'final_df' not in st.session_state:
    st.session_state.final_df = None

# 原始車隊人員名單 (含注音)
RAW_MEMBERS = [
    "丁秋吟(ㄉㄑㄧ)", "王永慶(ㄨㄩㄑ)", "王銓德(ㄨㄑㄉ)", "王志文(ㄨㄓㄨ)", "王家業(ㄨㄐㄧ)", "朱家樺(ㄓㄐㄏ)", "江旼珀(ㄐㄇㄆ)", "吳上苑(ㄨㄕㄩ)", "吳宜汶(ㄨㄧㄨ)", "呂恩昕(ㄌㄣㄒ)", 
    "呂淑惠(ㄌㄕㄏ)", "李國誥(ㄌㄍㄍ)", "李榮斌(ㄌㄖㄅ)", "李穎裕(ㄌㄧㄩ)", "阮智偉(ㄖㄓㄨ)", "周昆佑(ㄓㄎㄧ)", "周志暐(ㄓㄓㄨ)", "周志祥(ㄓㄓㄒ)", "林志嶸(ㄌㄓㄖ)", "林永松(ㄌㄩㄙ)", 
    "林志佳(ㄌㄓㄐ)", "劉柔君(ㄌㄖㄐ)", "林佳宏(ㄌㄐㄏ)", "林國芳(ㄌㄍㄈ)", "林婉茹(ㄌㄨㄖ)", "林智偉(ㄌㄓㄨ)", "林華盛(ㄌㄏㄕ)", "林瑞華(ㄌㄖㄏ)", "林瑞燿(ㄌㄖㄧ)", "林嘉信(ㄌㄐㄒ)", 
    "邱保銘(ㄑㄅㄇ)", "邱信培(ㄑㄒㄆ)", "徐琮凱(ㄒㄘㄎ)", "柯富強(ㄎㄈㄑ)", "洪偉倫(ㄏㄨㄌ)", "胡中興(ㄏㄓㄒ)", "胡耀仁(ㄏㄧㄖ)", "夏進通(ㄒㄐㄊ)", "涂旻聖(ㄊㄇㄕ)", "張進源(ㄓㄐㄩ)", 
    "張文男(ㄓㄨㄋ)", "張世明(ㄓㄕㄇ)", "張仕欣(ㄓㄕㄒ)", "張正中(ㄓㄓㄓ)", "張百江(ㄓㄅㄐ)", "張孟哲(ㄓㄇㄓ)", "張錦升(ㄓㄐㄕ)", "張勝富(ㄓㄕㄈ)", "張路雄(ㄓㄌㄒ)", "張鈞銘(ㄓㄐㄇ)", 
    "張仕宗(ㄓㄕㄗ)", "張聰捷(ㄓㄘㄐ)", "梁善鈞(ㄌㄕㄐ)", "郭政富(ㄍㄓㄈ)", "陳文政(ㄔㄨㄓ)", "陳仕明(ㄔㄕㄇ)", "陳怡如(ㄔㄧㄖ)", "陳瑜玲(ㄔㄩㄌ)", "陳冠良(ㄔㄍㄌ)", "陳春文(ㄔㄔㄨ)", 
    "陳敏訓(ㄔㄇㄒ)", "陳盛宏(ㄔㄕㄏ)", "陳進忠(ㄔㄐㄓ)", "陳佳忠(ㄔㄐㄓ)", "曾雅惠(ㄗㄧㄏ)", "游正豪(ㄧㄓㄏ)", "游育民(ㄧㄩㄇ)", "游振和(ㄧㄓㄏ)", "黃世政(ㄏㄕㄓ)", "黃明興(ㄏㄇㄒ)", 
    "黃信隆(ㄏㄒㄌ)", "黃冠霖(ㄏㄍㄌ)", "黃堃珉(ㄏㄎㄇ)", "黃賀進(ㄏㄏㄐ)", "黃麗萍(ㄏㄌㄆ)", "楊俊逸(ㄧㄐㄧ)", "魏捷祥(ㄨㄐㄒ)", "楊閔森(ㄧㄇㄙ)", "廖崇凱(ㄌㄔㄎ)", "廖竣傑(ㄌㄐㄐ)", 
    "廖致維(ㄌㄓㄨ)", "熊致堯(ㄒㄓㄧ)", "蔡芷纭(ㄘㄓㄩ)", "蔡玉雯(ㄘㄩㄨ)", "蔡榮峰(ㄘㄖㄈ)", "鄭銘宗(ㄓㄇㄗ)", "蕭大勛(ㄒㄉㄒ)", "蕭鈺潔(ㄒㄩㄐ)", "謝仁政(ㄒㄖㄓ)", "魏志龍(ㄨㄓㄌ)", 
    "藍慧真(ㄌㄏㄓ)", "詹獻睿(ㄓㄒㄖ)", "詹孟杰(ㄓㄇㄐ)", "曾舜麟(ㄗㄕㄌ)", "謝明佑(ㄒㄇㄧ)", "羅靜宜(ㄌㄐㄧ)", "姜小平(ㄐㄒㄆ)", "莊麗雪(ㄓㄌㄒ)", "吳慧芳(ㄨㄏㄈ)", "蔡瑞賓(ㄘㄖㄅ)", 
    "張上觀(ㄓㄕㄍ)", "陳素貞(ㄔㄙㄓ)", "吳建興(ㄨㄐㄒ)", "呂昇印(ㄌㄕㄧ)", "孟繁光(ㄇㄈㄍ)", "林坤茂(ㄌㄎㄇ)", "邱榮家(ㄑㄖㄐ)", "徐偉欽(ㄒㄨㄑ)", "鄭淵太(ㄓㄩㄊ)", "張志州(ㄓㄓㄓ)", 
    "張富山(ㄓㄈㄕ)", "陳珀升(ㄔㄆㄕ)", "陳溪宗(ㄔㄒㄗ)", "陳瑋楊(ㄔㄨㄧ)", "曾明雄(ㄗㄇㄒ)", "詹憲國(ㄓㄒㄍ)", "廖宏輝(ㄌㄏㄏ)", "廖翊均(ㄌㄧㄐ)", "劉邦杰(ㄌㄅㄐ)", "劉明煌(ㄌㄇㄏ)", 
    "蔡榮祥(ㄘㄖㄒ)", "蔡榮華(ㄘㄖㄏ)", "周智勤(ㄓㄓㄑ)", "謝志忠(ㄒㄓㄓ)", "涂欽耀(ㄊㄑㄧ)", "張志仲(ㄓㄓㄓ)", "蕭森巍(ㄒㄙㄨ)", "張銀恭(ㄓㄧㄍ)", "王正錄(ㄨㄓㄌ)", "曾建勳(ㄗㄐㄒ)", 
    "黃智煒(ㄏㄓㄨ)", "劉海森(ㄌㄏㄙ)", "賴南君(ㄌㄋㄐ)", "劉權漢(ㄌㄑㄏ)", "游伊君(ㄧㄧㄐ)", "陳勇志(ㄔㄩㄓ)", "陳文田(ㄔㄨㄊ)", "林秋雄(ㄌㄑㄒ)", "賴永昌(ㄌㄩㄔ)", "劉煉騰(ㄌㄌㄊ)", 
    "林錦志(ㄌㄐㄓ)", "林明忠(ㄌㄇㄓ)", "張哲誠(ㄓㄓㄔ)", "詹昆學(ㄓㄎㄒ)", "陳彥宏(ㄔㄧㄏ)", "許馨云(ㄒㄒㄩ)", "張橋語(ㄓㄑㄩ)"
]

# 重新格式化：依注音首字排序，確保相同首字的群聚在一起，並依照第二注音排序
ALL_MEMBERS_WITH_ZHUYIN = sorted([f"{m.split('(')[1][:-1]} - {m.split('(')[0]}" for m in RAW_MEMBERS])

# 乾淨的原始名單 (供產出報表用)
CLEAN_ALL_MEMBERS = [m.split('(')[0] for m in RAW_MEMBERS]

OPTIONS = ["--- 請點擊此處輸入注音或選擇人員 ---"] + ALL_MEMBERS_WITH_ZHUYIN

def generate_word_report(date_str, location_str, clean_attendees):
    doc = Document()
    doc.add_heading('木工機械單車協會 - 團騎點名紀錄', 0)
    doc.add_heading(f'日期：{date_str}', level=1)
    
    # 若有輸入地點，則將地點顯示在 Word 檔的標題下方
    if location_str:
        doc.add_heading(f'地點：{location_str}', level=2)
        
    doc.add_paragraph(f'本次參與總人數：{len(clean_attendees)} 人')
    doc.add_heading('出席名單：', level=2)
    doc.add_paragraph("、".join(clean_attendees))
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# 選定人員即簽到的連動函數
def on_person_select():
    selected = st.session_state.person_selector
    if selected != OPTIONS[0] and selected not in st.session_state.attendees:
        st.session_state.attendees.append(selected)
        st.session_state.report_generated = False # 有新人簽到，強制重新結算報表
    # 簽到後瞬間將選單切換回預設提示文字
    st.session_state.person_selector = OPTIONS[0]

st.title("🚴‍♂️ 車隊團騎點名系統")

# ==========================================
# 階段一：設定畫面 (完成後會自動隱藏)
# ==========================================
if not st.session_state.setup_complete:
    st.info("💡 請先完成下方設定。確認後介面會自動鎖定並進入點名模式。")
    
    st.markdown("### 步驟一：填寫活動資訊")
    selected_date = st.date_input("📅 點擊開啟月曆選擇日期：", datetime.today())
    # 新增地點輸入框
    event_location = st.text_input("📍 輸入本次團騎地點 (例如：鳳凰山、136縣道...)")
    
    st.markdown("### 步驟二：載入資料")
    mode = st.radio("請選擇資料載入方式：", ["上傳舊有總表接續點名", "建立全新總表"])
    
    temp_df = None
    if mode == "上傳舊有總表接續點名":
        uploaded_file = st.file_uploader("📂 選擇先前的 Excel 總表", type=["xlsx"])
        if uploaded_file is not None:
            temp_df = pd.read_excel(uploaded_file)
            st.success("✅ 舊表單載入成功！")
    else:
        temp_df = pd.DataFrame({"姓名": CLEAN_ALL_MEMBERS})
        st.info("🆕 將建立全新表單。")

    if st.button("🔒 確認設定並開始點名"):
        if mode == "上傳舊有總表接續點名" and temp_df is None:
            st.warning("⚠️ 請先上傳檔案再繼續！")
        else:
            st.session_state.df = temp_df
            st.session_state.date_str = selected_date.strftime("%Y-%m-%d")
            st.session_state.location = event_location.strip()
            st.session_state.setup_complete = True
            st.rerun() 

# ==========================================
# 階段二：點名主畫面 (雙欄位設計)
# ==========================================
else:
    # 顯示目前鎖定的日期與地點
    display_event = st.session_state.date_str
    if st.session_state.location:
        display_event += f" ({st.session_state.location})"
    st.success(f"📌 目前鎖定點名活動：**{display_event}**")
    
    col1, col2 = st.columns([1.5, 1])
    
    with col1:
        st.markdown("### 🔍 步驟三：快速簽到")
        st.write("💡 直接打注音首字 (如：打『ㄊ』優先找涂)，選到名字瞬間即完成簽到。")
        
        st.selectbox(
            "輸入注音或姓名關鍵字：", 
            OPTIONS, 
            key="person_selector", 
            on_change=on_person_select
        )
        
        st.markdown("---")
        st.markdown("#### 📝 修改/移除")
        updated_attendees = st.multiselect(
            "目前已簽到清單 (點擊 x 可移除誤點人員)：", 
            st.session_state.attendees, 
            default=st.session_state.attendees,
            label_visibility="collapsed"
        )
        if updated_attendees != st.session_state.attendees:
            st.session_state.attendees = updated_attendees
            st.session_state.report_generated = False
            st.rerun()

    with col2:
        st.markdown("### 📋 即時簽到清單")
        if not st.session_state.attendees:
            st.info("尚無人員簽到")
        else:
            for i, person in enumerate(st.session_state.attendees):
                clean_name = person.split(' - ')[1]
                st.write(f"**{i+1}.** {clean_name}")

    st.markdown("---")
    
    # ==========================================
    # 產出報表區塊 (按鈕推至右下角)
    # ==========================================
    col_empty1, col_empty2, col_btn = st.columns([2, 1, 2])
    with col_btn:
        finish_btn = st.button("💾 點名結束！", use_container_width=True)

    if finish_btn:
        df = st.session_state.df.copy() 
        date_str = st.session_state.date_str
        location_str = st.session_state.location
        
        # 將日期與地點組合，做為 Excel 欄位名稱
        event_col_name = f"{date_str} {location_str}".strip()
        
        if not st.session_state.attendees:
            st.error("⚠️ 目前沒有任何人簽到！")
        elif event_col_name in df.columns and not df[df[event_col_name] == 'V'].empty:
            st.error(f"⚠️ {event_col_name} 已經有點名紀錄囉！")
        else:
            clean_attendees = [p.split(' - ')[1] for p in st.session_state.attendees]
            
            missing_members = [m for m in CLEAN_ALL_MEMBERS if m not in df['姓名'].values]
            if missing_members:
                new_rows = pd.DataFrame({"姓名": missing_members})
                df = pd.concat([df, new_rows], ignore_index=True)

            if event_col_name not in df.columns:
                df[event_col_name] = ""
            for member in clean_attendees:
                df.loc[df['姓名'] == member, event_col_name] = "V"

            if '編號' in df.columns:
                df = df.drop(columns=['編號'])
            if '總次數' in df.columns:
                df = df.drop(columns=['總次數'])

            date_cols = [col for col in df.columns if col != '姓名']
            df['總次數'] = df[date_cols].apply(lambda x: (x == 'V').sum(), axis=1)

            df.insert(0, '編號', range(1, len(df) + 1))
            final_cols = ['編號', '姓名'] + date_cols + ['總次數']
            df = df[final_cols]
            
            st.session_state.final_df = df
            
            excel_buffer = BytesIO()
            df.to_excel(excel_buffer, index=False)
            st.session_state.excel_data = excel_buffer.getvalue()
            
            st.session_state.word_data = generate_word_report(date_str, location_str, clean_attendees)
            st.session_state.report_generated = True

    if st.session_state.report_generated:
        st.success(f"🎉 已成功記錄 {len(st.session_state.attendees)} 人！您可以隨時點擊下方按鈕下載檔案。")

        # 輸出檔案名稱也加上地點，方便日後整理
        file_name_suffix = st.session_state.date_str
        if st.session_state.location:
             file_name_suffix += f"_{st.session_state.location}"
             
        dl_col1, dl_col2 = st.columns(2)
        with dl_col1:
            st.download_button("📥 下載 Excel 總表", data=st.session_state.excel_data, file_name=f"車隊點名總表_{file_name_suffix}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with dl_col2:
            st.download_button("📥 下載 Word 紀錄檔", data=st.session_state.word_data, file_name=f"團騎點名紀錄_{file_name_suffix}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        
        st.write("📊 總表預覽：")
        st.dataframe(st.session_state.final_df)

    st.markdown("---")
    if st.button("⚙️ 重新設定日期或表單"):
        st.session_state.setup_complete = False
        st.rerun()
