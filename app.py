import streamlit as st
import tempfile
import os
import sys
from pathlib import Path

# 将当前目录加入 sys.path，以便导入 thesis_formatter 模块
sys.path.insert(0, str(Path(__file__).parent))
from thesis_formatter import format_thesis

st.set_page_config(page_title="论文格式一键排版工具", page_icon="📄")
st.title("📄 论文格式一键排版工具")
st.markdown("上传 Word 文档，自动按照模板调整页边距、标题样式、页眉页脚、页码等。")

# 侧边栏参数
with st.sidebar:
    st.header("⚙️ 参数设置")

    # 排版模式选择
    mode = st.radio(
        "排版模式",
        options=["xsyu_thesis", "homework"],
        format_func=lambda x: "🎓 西安石油大学毕业论文模式" if x == "xsyu_thesis" else "📝 日常大作业模式",
        index=0
    )

    # 根据模式显示不同的输入项
    if mode == "homework":
        course_title = st.text_input("课程名称（页眉文字）", value="课程作业")
        doc_title = course_title
    else:
        doc_title = st.text_input("论文标题（页眉文字）", value="论文")
        course_title = None

    st.markdown("---")
    st.markdown("**📋 模式说明**：")

    if mode == "xsyu_thesis":
        st.markdown("- 封面+摘要：罗马数字页码（Ⅰ, Ⅱ, Ⅲ...）")
        st.markdown("- 正文各章：独立节，节间自动分节")
        st.markdown("- 奇数页页眉：**章节标题**")
        st.markdown("- 偶数页页眉：**西安石油大学本科毕业设计(论文)**")
        st.markdown("- 正文页码从1开始（阿拉伯数字）")
        st.markdown("- 自动识别：第一章/一、/1. /摘要 等标题格式")
    else:
        st.markdown("- 无章节分节符，文档为单一节")
        st.markdown("- 页眉为课程名称，居中显示")
        st.markdown("- 连续页码，从第1页开始")
        st.markdown("- 无奇偶页区分")
        st.markdown("- 自动识别各类标题格式")

    st.markdown("---")
    st.markdown("**📌 通用说明**：")
    st.markdown("- 仅支持 `.docx` 格式")
    st.markdown("- 自动处理封面、摘要、正文、参考文献样式")
    st.markdown("- 自动清除 Markdown 残留标记")
    st.markdown("- 章节之间不插入分页符（连续排列）")

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
                format_thesis(
                    input_path=input_path,
                    output_path=output_path,
                    doc_title=doc_title,
                    no_toc=False,
                    mode=mode,
                    course_title=course_title if mode == "homework" else None
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
    st.info("👆 请先上传 .docx 文件，然后点击「开始格式化」")
