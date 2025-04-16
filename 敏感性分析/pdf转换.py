import asyncio
from pathlib import Path
from kreuzberg import extract_file, extract_bytes

# Basic file extraction
async def extract_document():
    # 图片文件路径
    image_path = r"D:\小工具\cursor\图片\scan.png"  # 使用原始字符串或双反斜杠
    
    # 从图片提取文本
    img_result = await extract_file(image_path)
    print(f"图片文本: {img_result.content}")

# 处理上传的文件（如果是字节内容上传）
async def process_upload(file_content: bytes, mime_type: str):
    """Process uploaded file content with known MIME type."""
    result = await extract_bytes(file_content, mime_type=mime_type)
    return result.content

# 示例：不同类型文件的使用
async def handle_uploads(pdf_bytes, image_bytes, docx_bytes):
    # 处理PDF文件上传
    pdf_result = await extract_bytes(pdf_bytes, mime_type="application/pdf")
    print(f"PDF内容: {pdf_result}")
    
    # 处理图片文件上传
    img_result = await extract_bytes(image_bytes, mime_type="image/jpeg")
    print(f"图片内容: {img_result}")
    
    # 处理Word文档文件上传
    docx_result = await extract_bytes(docx_bytes, mime_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    print(f"Word内容: {docx_result}")

# 运行所有异步函数
async def main():
    # 提供图片文件路径或字节数据
    await extract_document()

# 启动事件循环并运行异步任务
if __name__ == "__main__":
    asyncio.run(main())
