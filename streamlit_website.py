import streamlit as st
import pandas as pd
from docx import Document

# Excel转TXT函数
def excel_to_txt():
   # 添加列名输入框
    column_name_for_questions = st.text_input("请输入问题列的列名:", "")
    column_name_for_answers = st.text_input("请输入答案列的列名:", "")

    # 添加一个上传文件的按钮
    uploaded_file = st.file_uploader("请选择一个Excel文件", accept_multiple_files=False)

    # 检查是否有文件被上传
    if uploaded_file is not None:
        # 读取Excel文件
        try:
            df = pd.read_excel(uploaded_file)
            
            # 检查用户输入的列名是否存在于DataFrame中
            if column_name_for_questions in df.columns and column_name_for_answers in df.columns:
                # 提示用户文件已上传
                st.write("文件已成功上传。")

                # 提取用户指定的列
                df_questions = df[column_name_for_questions]
                df_answers = df[column_name_for_answers]

                # 处理文件
                # 在问题列前面加上 "问题："，在答案列前面加入 "答案："
                df['问题'] = '问题：' + df_questions.astype(str)
                df['答案'] = '答案：' + df_answers.astype(str)

                # 将处理后的数据转换为文本格式
                output_text = ""
                for index, row in df.iterrows():
                    output_text += row['问题'] + '\n'
                    output_text += row['答案'] + '\n'
                    output_text += '==\n'

                # 提供下载处理后文件的选项
                download_link = st.download_button(
                    label="下载处理后的文件",
                    data=output_text.encode('utf-8'),
                    file_name='output.txt',
                    mime='text/plain'
                )

                # 如果下载链接被点击，显示一个消息
                if download_link:
                    st.write("文件已准备好下载。")
            else:
                st.error("列名不存在于Excel文件中，请检查列名是否正确。")
        except Exception as e:
            st.error(f"读取文件时发生错误：{e}")
    else:
        st.write("请上传一个Excel文件。")

# 切片统计函数
def slice_statistics(file_type, uploaded_file):
    if uploaded_file is not None:
        if file_type == 'txt':
            # 读取txt文件
            text = uploaded_file.read().decode('utf-8')
            blocks = text.split("==")
            text_blocks = [block.strip() for block in blocks if block.strip() != ""]
            # 移除空白行和仅包含空白的段落
            # 计算每个文本块的字数
            block_lengths = [len(block) for block in text_blocks]
            # 保存到DataFrame
            df = pd.DataFrame({"Text Block": text_blocks, "Length": block_lengths})
            # 找出字数最多的段落
            max_length = max(block_lengths)
            max_length_block = text_blocks[block_lengths.index(max_length)]
            # 保存到xlsx文件
            output_filename = 'text_blocks_lengths.xlsx'
            df.to_excel(output_filename, index=False)
            
            # 显示下载按钮
            with open(output_filename, "rb") as file:
                btn = st.download_button(
                    label=f"下载 {output_filename}",
                    data=file.read(),
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            st.write(f"TXT文件切片统计完成，结果已保存到 {output_filename}")
            st.write(f"字数最多的段落是：'{max_length_block}'，包含字数：{max_length}")

        elif file_type == 'word':
            # 获取文件名和类型
            file_name = uploaded_file.name
            file_ext = file_name.split('.')[-1]

            if file_ext == 'docx':
                # 使用python-docx读取Word文档
                doc = Document(uploaded_file)
                content = []
                for para in doc.paragraphs:
                    content.append(para.text)
                text = '=='.join(content)
                slices = text.split('==')
                slices = [s.strip() for s in slices if s.strip()]  # 移除空切片

                # 计算每个切片的字数长度
                slice_lengths = [(s, len(s)) for s in slices]

                # 创建DataFrame
                df = pd.DataFrame(slice_lengths, columns=['切片', '字数'])

                # 找出字数最多的段落
                max_length = df['字数'].max()
                max_length_slice = df.loc[df['字数'] == max_length, '切片'].values[0]

                # 保存到Excel文件
                output_path = 'slice_lengths.xlsx'
                df.to_excel(output_path, index=False, encoding='utf-8')
                
                # 显示下载按钮
                with open(output_path, "rb") as file:
                    btn = st.download_button(
                        label=f"下载 {output_path}",
                        data=file.read(),
                        file_name=output_path,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                st.write(f"Word文件切片统计完成，结果已保存到 {output_path}。")
                st.write(f"字数最多的切片是：'{max_length_slice}'，包含字数：{max_length}")
            else:
                st.write("暂不支持该文件格式，请上传 .docx 文件。")

# 创建Streamlit应用
st.title('文件处理工具')

# 添加标签页
tabs = st.tabs(['Excel转TXT', '切片统计'])
tab1, tab2 = tabs[0], tabs[1]

with tab1:
    st.header('Excel转TXT')
    excel_to_txt()

with tab2:
    st.header('切片统计')
    file_type = st.selectbox(
        '选择文件类型',
        ('word', 'txt')
    )
    uploaded_file = st.file_uploader("请选择要统计的文件", type=["docx", "txt"])
    if uploaded_file is not None and file_type in ['word', 'txt']:
        slice_statistics(file_type, uploaded_file)
