# 文件处理工具

这是一个使用Python和Streamlit构建的Web应用程序，用于处理文件并提供两种主要功能：Excel转TXT和切片统计。

## 功能

### Excel转TXT
- 用户可以上传Excel文件，并指定问题和答案的列名。
- 程序将读取Excel文件，提取指定列，并在每行前分别添加“问题：”和“答案：”标签。
- 最终，程序将转换后的数据提供为文本文件下载。

### 切片统计
- 用户可以选择上传txt或docx文件，并进行切片统计。
- 对于txt文件，程序将统计每个文本块的字数，并找出字数最多的段落。
- 对于Word文档，程序将统计每个段落的字数，并找出字数最多的段落。
- 统计结果将保存为Excel文件，并提供下载。

## 快速开始

### 环境要求
- Python 3.x
- Streamlit
- Pandas
- python-docx (仅Word文档处理功能需要)

### 安装步骤
1. 克隆仓库到本地机器
   ```bash
   git clone https://github.com/dinggithub/qa_excel2txt/.git
   
2. 进入项目目录
   ```bash
   cd yourrepository
   
3. 创建虚拟环境（可选）
   ```bash
   python -m venv venv

4. 激活虚拟环境

Windows：
.\venv\Scripts\activate

macOS/Linux：
source venv/bin/activate

5. 安装依赖
   ```bash
   pip install -r requirements.txt

### 运行应用
在项目根目录下运行以下命令启动Web应用：
   ```bash
      streamlit run streamlit_website.py

然后在浏览器中打开 http://localhost:8501 来访问应用。
