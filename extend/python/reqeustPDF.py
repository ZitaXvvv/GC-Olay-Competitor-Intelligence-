import requests



def RequestPDF(url):
    save_path = r'C:\Users\sun.y.32\OneDrive - Procter and Gamble\Desktop\RPA\美妆网站项目\PDF\file.pdf'
    response = requests.get(url)
    response.raise_for_status()  # 检查请求是否成功

    with open(save_path, "wb") as file:
        file.write(response.content)

    print("文件下载完成")



RequestPDF('https://hzpba.nmpa.gov.cn/HZPBZCX/PTHZPBA-SERVER/nmpafile/gsxxFilePreview?attachmentId=1233094695292313600') 

