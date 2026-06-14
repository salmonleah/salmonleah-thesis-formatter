import streamlit as st
import tempfile
import os
import sys
from pathlib import Path

# 将当前目录加入 sys.path，以便导入 thesis_formatter 模块
sys.path.insert(0, str(Path(__file__).parent))
from thesis_formatter import format_thesis  # 修正：原为 format_document

st.set_page_config(page_title="论文格式一键排版工具", page_icon="📄")
st.title("📄 西安石油大学论文格式一键排版")
st.markdown("上传 Word 文档，自动按照学校模板调整页边距、标题样式、页眉页脚、页码等。")

# 侧边栏参数
with st.sidebar:
    st.header("⚙️ 参数设置")
    doc_title = st.text_input("文档标题（页眉文字）", value="论文")
    st.markdown("---")
    st.markdown("**说明**：")
    st.markdown("- 仅支持 `.docx` 格式")
    st.markdown("- 自动处理封面、摘要、正文、参考文献样式")
    st.markdown("- 正文部分不会插入分页符，章节连续排列")
    st.markdown("- 处理完成后可下载新文档")

# 主区域：文件上传
uploaded_file = st.file_uploader("选择 Word 文档 (.docx)", type=["docx"])

if uploaded_file is not None:
    # 保存上传的临时文件
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_input:
        tmp_input.write(uploaded_file.read())
        input_path = tmp_input.name

    # 输出文件路径（在临时目录中生成）
    output_path = input_path.replace(".docx", "_formatted.docx")

    if st.button("✨ 开始格式化", type="primary"):
        with st.spinner("正在处理，请稍候..."):
            try:
                # 调用核心函数，修正为 format_thesis
                format_thesis(
                    input_path=input_path,
                    output_path=output_path,
                    doc_title=doc_title,
                    no_toc=False   # 保留目录提示，不自动生成
                )
                st.success("✅ 格式化完成！")

                # 读取结果文件并提供下载
                with open(output_path, "rb") as f:
                    st.download_button(
                        label="📥 下载格式化后的文档",
                        data=f,
                        file_name=f"formatted_{uploaded_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            except Exception as e:
                st.error(f"处理出错：{e}")
            finally:
                # 清理临时文件
                try:
                    os.unlink(input_path)
                    if os.path.exists(output_path):
                        os.unlink(output_path)
                except:
                    pass
else:
    st.info("👆 请先上传 .docx 文件")