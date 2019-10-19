from PIL import Image
from PIL import ImageChops 

def compare_images(path_one, path_two, diff_save_location):
    """
    比较图片，如果有不同则生成展示不同的图片
 
    @参数一: path_one: 第一张图片的路径
    @参数二: path_two: 第二张图片的路径
    @参数三: diff_save_location: 不同图的保存路径
    """
    image_one = Image.open(path_one)
    image_two = Image.open(path_two)
 
    diff = ImageChops.difference(image_one, image_two)
 
    if diff.getbbox() is None:
        # 图片间没有任何不同则直接退出
        return
    else:
        diff.save(diff_save_location)
 
if __name__ == '__main__':
    compare_images(r'D:\测试中心\20190826金龙鱼稻谷鲜生湖畔农家香米5KG奥运版2奇之意虹（镂空）.jpg',
                   r'D:\测试中心\20190906金龙鱼稻谷鲜生湖畔农家香米5KG奥运版3奇之意虹（镂空）-作废.jpg',
                   r'D:\测试中心\diff.jpg')
