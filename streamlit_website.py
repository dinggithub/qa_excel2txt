import streamlit as st
import pandas as pd

# 创建一个标题
st.title('Excel 文件处理和下载')

# 添加列名输入框
column_name_for_questions = st.text_input("请输入问题列的列名:", "问题")
column_name_for_answers = st.text_input("请输入答案列的列名:", "答案")

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