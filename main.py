import json
import os
import re
import requests
from io import BytesIO
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm

def run(inputs: dict) -> dict:
    """
    Dify 插件标准入口函数
    """
    # 1. 获取输入参数
    template_file_info = inputs.get("template_file", {})
    template_path = template_file_info.get("path") 
    fill_data_str = inputs.get("fill_data", "{}")
    img_width_cm = inputs.get("default_image_width", 12)

    if not template_path or not os.path.exists(template_path):
        raise FileNotFoundError("【插件错误】未能正确读取到上传的 Word 模板文件。")

    # 2. 解析 JSON 数据
    try:
        # 兼容处理：防止传入带 Markdown 代码块包装的 JSON
        if "```json" in fill_data_str:
            fill_data_str = fill_data_str.split("```json")[1].split("```")[0].strip()
        elif "```" in fill_data_str:
            fill_data_str = fill_data_str.split("```")[1].strip()
        
        context = json.loads(fill_data_str.strip())
    except Exception as e:
        raise ValueError(f"【插件错误】输入的填充数据不是合法的 JSON 格式，请检查上游节点输出。错误详情: {str(e)}")

    # 3. 初始化 Word 模板
    try:
        doc = DocxTemplate(template_path)
    except Exception as e:
        raise RuntimeError(f"【插件错误】打开 Word 模板失败，请确保上传的是标准 .docx 格式文件: {str(e)}")

    # 4. 核心处理：遍历 JSON 上下文，智能拦截并处理图片 URL
    processed_context = {}
    
    # 匹配图片 URL 的正则表达式 (支持带有查询参数的常见图片格式)
    img_url_pattern = re.compile(r'^https?://.*\\.(?:png|jpg|jpeg|gif|bmp|webp)(?:\\?.*)?$', re.IGNORECASE)
    # 额外兼容：Dify 内部流转的文件预览 URL 形如 http://.../file-preview?timestamp=...
    dify_file_pattern = re.compile(r'^https?://.*/file-preview\\?.*$', re.IGNORECASE)

    for key, value in context.items():
        val_str = str(value).strip()
        
        # 判断该字段是否为图片网络链接
        if img_url_pattern.match(val_str) or dify_file_pattern.match(val_str):
            try:
                # 顺着网络请求下载图片到内存
                response = requests.get(val_str, timeout=15)
                if response.status_code == 200:
                    image_stream = BytesIO(response.content)
                    # 将图片转化为 docxtpl 专用的 InlineImage 嵌入对象，并设定厘米宽度
                    processed_context[key] = InlineImage(doc, image_stream, width=Cm(img_width_cm))
                else:
                    # 下载失败时，保留原始 URL 字符串，防止程序崩溃
                    processed_context[key] = f"[图片下载失败，状态码 {response.status_code}]: {val_str}"
            except Exception as e:
                processed_context[key] = f"[图片下载异常]: {val_str}，错误: {str(e)}"
        else:
            # 普通文本直接透传
            processed_context[key] = value

    # 5. 执行渲染模板填空
    try:
        doc.render(processed_context)
        
        # 6. 将生成的成品 Word 存放到 Dify 沙箱指定的临时目录下
        output_filename = "申报书_精准填充成品.docx"
        output_path = os.path.join(os.path.dirname(template_path), output_filename)
        doc.save(output_path)
    except Exception as e:
        raise RuntimeError(f"【插件错误】Word 模板标签替换渲染失败，请检查模板中的双大括号 {{}} 语法是否闭合。详情: {str(e)}")

    # 7. 按照 Dify 插件规范返回文件变量
    return {
        "result_file": {
            "path": output_path,
            "name": output_filename,
            "mime_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        }
    }
