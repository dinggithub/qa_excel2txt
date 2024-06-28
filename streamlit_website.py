import streamlit as st
import pandas as pd
import docx
import PyPDF2
from io import BytesIO
from datetime import datetime
import os
import chardet

def excel_to_txt():
    st.header("Excel转QA对TXT")
    uploaded_file = st.file_uploader("上传Excel文件，将问题列和答案列生成QA对TXT文件", type=["xlsx"])
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

def extract_pdf():

    st.header("提取PDF页面")

    # 上传PDF文件
    uploaded_file = st.file_uploader("选择要处理的PDF文件，可将页面提取为PDF文件", type="pdf")

    if uploaded_file is not None:
        # 获取用户输入的页面范围
        start_page = st.number_input("开始页码", min_value=1, value=1, step=1)
        end_page = st.number_input("结束页码", min_value=1, value=1, step=1)

        if st.button("提取页面"):
            # 读取PDF文件
            pdf_reader = PyPDF2.PdfReader(uploaded_file)
            num_pages = len(pdf_reader.pages)

            # 检查页码范围是否合法
            if start_page < 1 or end_page < start_page or end_page > num_pages:
                st.error("无效的页码范围,请重新输入。")
            else:
                # 提取指定页面范围并保存为新的PDF文件
                for page_num in range(start_page - 1, end_page):
                    new_pdf = PyPDF2.PdfWriter()
                    new_pdf.add_page(pdf_reader.pages[page_num])

                    new_pdf_filename = f"page_{page_num + 1}.pdf"
                    with open(new_pdf_filename, "wb") as new_file:
                        new_pdf.write(new_file)

                    # 将新的PDF文件作为下载链接提供
                    with open(new_pdf_filename, "rb") as file:
                        st.download_button(
                            label=f"下载 {new_pdf_filename}",
                            data=file.read(),
                            file_name=new_pdf_filename,
                            mime="application/pdf",
                        )

                st.success("页面提取完成!")

def detect_encoding(file):
    raw_data = file.read()
    result = chardet.detect(raw_data)
    encoding = result['encoding']
    file.seek(0)  # 重置文件指针
    return encoding

def merge_txt_files():
    uploaded_files = st.file_uploader("上传多个TXT文件", type="txt", accept_multiple_files=True)
    
    if uploaded_files:
        uploaded_files.sort(key=lambda x: x.name)
        merged_content = ""
        
        for uploaded_file in uploaded_files:
            encoding = detect_encoding(uploaded_file)
            if encoding is None:
                encoding = 'utf-8'  # 默认使用utf-8编码
            
            try:
                content = uploaded_file.read().decode(encoding)
            except UnicodeDecodeError:
                # 尝试使用其他常见编码
                try:
                    content = uploaded_file.read().decode('gb2312')
                except UnicodeDecodeError:
                    st.error(f"文件 {uploaded_file.name} 解码失败，无法处理该文件。")
                    return
            
            merged_content += content + "\n"
        
        merged_file_name = "merged_file.txt"
        with open(merged_file_name, "w", encoding="utf-8") as merged_file:
            merged_file.write(merged_content)
        
        with open(merged_file_name, "rb") as merged_file:
            st.download_button(
                label="下载合并后的TXT文件",
                data=merged_file,
                file_name=merged_file_name,
                mime="text/plain"
            )

st.title("文件处理工具")
tabs = st.tabs(["Excel转TXT", "切片统计", "PDF转TXT", "提取PDF页面", "合并TXT文件"])

with tabs[0]:
    excel_to_txt()

with tabs[1]:
    slice_statistics()

with tabs[2]:
    pdf_to_txt()

with tabs[3]:
    extract_pdf()
    
with tabs[4]:
    merge_txt_files()
