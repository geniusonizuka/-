import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
from io import BytesIO

# 車隊人員名單 (隊員與顧問)
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

# 1. 雲端版專屬：上傳舊檔案接續紀錄
st.markdown("### 步驟一：載入資料")
uploaded_file = st.file_uploader("📂 選擇之前的 Excel 總表 (若是第一次點名請略過此步驟)", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.success("舊表單載入成功！可接續點名。")
else:
    df = pd.DataFrame({"姓名": ALL_MEMBERS, "總次數": 0})
    st.info("目前為全新表單。")

# 2. 選擇日期與出席人員
st.markdown("### 步驟二：今日點名")
selected_date = st.date_input("📅 選擇團騎日期", datetime.today())
date_str = selected_date.strftime("%Y-%m-%d")

attendees = st.multiselect("✅ 請勾選今日出席人員 (可輸入姓名搜尋)：", ALL_MEMBERS)

# 3. 執行點名與輸出
if st.button("💾 完成點名並產出報表"):
    if not attendees:
        st.warning("請至少選擇一位出席人員！")
    elif date_str in df.columns and not df[df[date_str] == 'V'].empty:
        st.warning(f"{date_str} 已經有紀錄囉，請確認日期是否正確。")
    else:
        # 確保名單完整
        missing_members = [m for m in ALL_MEMBERS if m not in df['姓名'].values]
        if missing_members:
            new_rows = pd.DataFrame({"姓名": missing_members, "總次數": 0})
            df = pd.concat([df, new_rows], ignore_index=True)

        if date_str not in df.columns:
            df[date_str] = ""

        # 寫入打勾標記 (V)
        for member in attendees:
            df.loc[df['姓名'] == member, date_str] = "V"
            
        # 重新計算總次數
        date_columns = [col for col in df.columns if col not in ['姓名', '總次數']]
        df['總次數'] = df[date_columns].apply(lambda x: (x == 'V').sum(), axis=1)

        st.success(f"已成功記錄 {len(attendees)} 人！請務必下載下方更新後的檔案。")

        excel_buffer = BytesIO()
        df.to_excel(excel_buffer, index=False)
        word_bytes = generate_word_report(date_str, attendees)

        st.markdown("### 步驟三：下載今日檔案")
        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                label="📥 下載更新後的 Excel 總表",
                data=excel_buffer.getvalue(),
                file_name=f"車隊點名總表_{date_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col2:
            st.download_button(
                label="📥 下載當日 Word 紀錄檔",
                data=word_bytes,
                file_name=f"{date_str}_團騎點名紀錄.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        
        st.write("目前總表預覽：")
        st.dataframe(df)
