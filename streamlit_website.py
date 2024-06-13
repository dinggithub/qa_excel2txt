import streamlit as st
import pandas as pd
import docx
import PyPDF2
from io import BytesIO
from datetime import datetime

def excel_to_txt():
    st.header("Excel转TXT")
    uploaded_file = st.file_uploader("上传Excel文件", type=["xlsx"])
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        st.write("Excel文件内容：", df)
        
        question_col = st.text_input("输入问题列的列名")
        answer_col = st.text_input("输入答案列的列名")
        
        if st.button("生成TXT文件"):
            if question_col in df.columns and answer_col in df.columns:
                output = ""
                for index, row in df.iterrows():
                    output += f"问题：{row[question_col]}\n答案：{row[answer_col]}\n==\n"
                
                st.download_button(
                    label="下载TXT文件",
                    data=output.encode('utf-8'),
                    file_name="output.txt",
                    mime="text/plain"
                )
            else:
                st.error("列名不正确，请检查输入的列名。")

def slice_statistics():
    st.header("切片统计")
    
    # 使用session_state来跟踪文件类型的变化
    if "file_type" not in st.session_state:
        st.session_state.file_type = "TXT"
    
    file_type = st.selectbox("选择文件类型", ["TXT", "Word (.docx)"], index=0 if st.session_state.file_type == "TXT" else 1)
    
    # 如果文件类型发生变化，清除上传的文件
    if file_type != st.session_state.file_type:
        st.session_state.file_type = file_type
        st.session_state.uploaded_file = None
    
    # 使用st.empty()来重新渲染文件上传组件
    file_uploader_placeholder = st.empty()
    uploaded_file = file_uploader_placeholder.file_uploader("上传文件", type=["txt", "docx"])
    
    if uploaded_file:
        try:
            if file_type == "TXT":
                content = uploaded_file.read().decode('utf-8')
                blocks = content.split("==")
            elif file_type == "Word (.docx)":
                # 确保文件类型正确
                if uploaded_file.name.endswith(".docx"):
                    doc = docx.Document(uploaded_file)
                    content = "\n".join([para.text for para in doc.paragraphs if para.text.strip() != ""])
                    blocks = content.split("==")
                else:
                    st.error("上传的文件不是有效的Word文件，请重新上传。")
                    return
            
            block_lengths = [len(block) for block in blocks]
            max_length = max(block_lengths)
            max_block = blocks[block_lengths.index(max_length)]
            
            st.write(f"字数最多的文本块/段落：{max_block}")
            st.write(f"字数：{max_length}")
            
            if st.button("生成Excel文件"):
                output = BytesIO()
                writer = pd.ExcelWriter(output, engine='xlsxwriter')
                df = pd.DataFrame({"文本块/段落": blocks, "字数": block_lengths})
                df.to_excel(writer, index=False)
                writer.close()  # 使用close()方法保存文件
                output.seek(0)
                
                st.download_button(
                    label="下载Excel文件",
                    data=output,
                    file_name="output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"文件处理时出错: {e}")

def pdf_to_txt():
    st.header("PDF转TXT")
    uploaded_file = st.file_uploader("上传PDF文件", type=["pdf"])
    if uploaded_file:
        reader = PyPDF2.PdfReader(uploaded_file)
        output = ""
        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]
            output += page.extract_text() + "\n"
        
        st.download_button(
            label="下载TXT文件",
            data=output.encode('utf-8'),
            file_name="output.txt",
            mime="text/plain"
        )

st.title("文件处理工具")
tabs = st.tabs(["Excel转TXT", "切片统计", "PDF转TXT"])

with tabs[0]:
    excel_to_txt()

with tabs[1]:
    slice_statistics()

with tabs[2]:
    pdf_to_txt()
