import requests
import time

# PDF文件的URL
url = "https://www.95598.cn/omg-static/99302251957052709409300507214865.pdf"

# 添加请求头
headers = {
         "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
        'Accept':'application/pdf,*/*',
        'Accept-Encoding': 'gzip, deflate, br',
        'Connection': 'keep-alive',
         'Referer': 'https://95598.csg.cn/'
        }

for attempt in range(5):  # 尝试5次
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        # 保存PDF文件
        with open("电价信息.pdf", "wb") as f:
            f.write(response.content)
        print("PDF下载成功，已保存为 '电价信息.pdf'")
        break
    else:
        print(f"下载失败，状态码: {response.status_code}，正在重试...")
        time.sleep(2)  # 等待2秒后重试

                # 使用 save_pdf 下载PDF
        if 'url' in selected_subproject:
            print(f"准备下载PDF，URL: {selected_subproject['url']}")
            save_pdf(selected_subproject['url'], f"{selected_subproject['text']}_电价标准.pdf")
        else:
            print("未找到PDF的URL")
