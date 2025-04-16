import fitz  # PyMuPDF

# 打开PDF文件
pdf_document = fitz.open(r"D:\小工具\cursor\downloads\云南电网有限责任公司2025年02月代理购电工商业用户电价表 (2025-01-24)V4.pdf")
# 获取PDF的页数
num_pages = pdf_document.page_count

# 提取页面图像并保存为PNG文件
for page_num in range(len(pdf_document)):
    page = pdf_document.load_page(page_num)
    image_list = page.get_images(full=True)
    for img_index, img in enumerate(image_list):
        xref = img[0]
        base_image = pdf_document.extract_image(xref)
        image_bytes = base_image["image"]
        with open(f"extracted_image_{page_num}_{img_index}.png", "wb") as img_file:
            img_file.write(image_bytes)

# 可以使用openpyxl将提取的图像插入Excel