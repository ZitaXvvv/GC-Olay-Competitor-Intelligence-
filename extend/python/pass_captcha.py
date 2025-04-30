import requests
from hashlib import md5

class ChaojiyingClient(object):

    def __init__(self, username, password, soft_id):
        self.username = username
        password =  password.encode('utf8')
        self.password = md5(password).hexdigest()
        self.soft_id = soft_id
        self.base_params = {
            'user': self.username,
            'pass2': self.password,
            'softid': self.soft_id,
        }
        self.headers = {
            'Connection': 'Keep-Alive',
            'User-Agent': 'Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 5.1; Trident/4.0)',
        }

    def PostPic(self, im, codetype):
        """
        im: 图片字节
        codetype: 题目类型 参考 http://www.chaojiying.com/price.html
        """
        params = {
            'codetype': codetype,
        }
        params.update(self.base_params)
        files = {'userfile': ('ccc.jpg', im)}
        r = requests.post('http://upload.chaojiying.net/Upload/Processing.php', data=params, files=files, headers=self.headers)
        return r.json()

def slider(pic_path: str) -> int:
    """传入图片路径，返回滑块移动距离"""
    username = '18028585938'
    password = 'Cisco,123!'
    soft_id = '928592'
    client = ChaojiyingClient(username, password, soft_id)
    with open(pic_path, 'rb') as f:
        im = f.read()
        res = client.PostPic(im, 9102)['pic_str']
    x1_s = res.split('|')[0].split(',')[0]
    x2_s = res.split('|')[1].split(',')[0]
    x1, x2 = int(x1_s), int(x2_s)
    offset = abs(x1-x2)

    return offset


if __name__ == '__main__':
    print(slider('slider.png'))