import streamlit as st
import tempfile
import os
from thesis_formatter import format_document

st.set_page_config(page_title="论文排版工具", page_icon="📄")
st.title("📄 论文排版一键格式化工具")
st.markdown("上传 Word 文档，自动按照西安石油大学本科毕业论文模板进行格式化。")

uploaded_file = st.file_uploader("选择 Word 文档（.docx）", type=["docx"])

if uploaded_file is not None:
    # 保存临时文件
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_input:
        tmp_input.write(uploaded_file.read())
        input_path = tmp_input.name

    if st.button("开始格式化"):
        with st.spinner("正在处理，请稍候..."):
            try:
                output_path = format_document(input_path)
                with open(output_path, "rb") as f:
                    st.download_button(
                        label="📥 下载格式化后的文档",
                        data=f,
                        file_name="formatted_thesis.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                st.success("格式化完成！")
            except Exception as e:
                st.error(f"处理出错：{e}")
            finally:
                # 清理临时文件
                os.unlink(input_path)
                if os.path.exists(output_path):
                    os.unlink(output_path)