import cv2
import numpy as np

# замена одного элемента на другой в строке
def change_elements(string, element_old, element_new):
    string = list(string)
    for i in range(len(string)):
        if string[i]==element_old:
            string[i] = element_new
    string = ''.join(string)
    return string


# cv2 сохранение изо с символами не из ascii
def cv2_imencode(img_path_to, img):
    success, im_buf_arr = cv2.imencode('.jpg', img)
    if success:
        im_buf_arr.tofile(img_path_to)


# извлечение цифр из строки
def extract_text_number(string):
    number = ''
    try:
        for s in string:
            if s.isdigit():
                number += s
        return int(number)
    except:
        return ''


# исправление артикля (возможны добавки Арт. и др символы)
def correct_article(string):
    try:
        return ''.join([s for s in string if s.isdigit() or 65 <= ord(s) <= 90 or s == '-'])
    except:
        pass

# проверка есть среди артикулов в экселе какой-то в названии файла или папки
def check_item(unique_items, item_photo):
    for unique_item in unique_items:
        elements = unique_item.split('-')
        max_len = max([len(element) for element in elements])
        for element in elements:
            if len(element)==max_len and element in item_photo:
                return unique_item

# выравнивает изображение с прямыми линиями до горизонтали или вертикали (куда ближе)
def horiz(image):
    canny = cv2.Canny(image, 50, 150)
    lines = cv2.HoughLinesP(canny, 1, np.pi/180, threshold=100, minLineLength=100, maxLineGap=100)
    # Вычисление угла поворота
    angle_h_1 = []
    angle_h_2 = []
    angle_v_1 = []
    angle_v_2 = []
    if lines is not None:
        for line in lines:
            x1, y1, x2, y2 = line[0]
            ang = np.degrees(np.arctan2(y2 - y1, x2 - x1))
            # print(ang)
            if 0 <= ang < 45:
                # cv2.line(image,(x1, y1), (x2, y2),(0,0,255),15)
                angle_h_1.append(ang)

            if -45 < ang < 0:
                # cv2.line(image,(x1, y1), (x2, y2),(255,0,0),15)
                angle_h_2.append(ang)

            if 45 <= ang < 90:
                # cv2.line(image,(x1, y1), (x2, y2),(0,255,0),15)
                angle_v_1.append(ang)
            if -90 < ang < -45:
                # cv2.line(image,(x1, y1), (x2, y2),(0,255,0),15)
                angle_v_2.append(ang)

    angle_h_1 = np.array(angle_h_1, dtype = 'float32')
    angle_v_1 = np.array(angle_v_1, dtype = 'float32')
    angle_h_2 = np.array(angle_h_2, dtype = 'float32')
    angle_v_2 = np.array(angle_v_2, dtype = 'float32')
    # print('angle_h_1', angle_h_1)
    # print('angle_v_1', angle_v_1)
    # print('angle_h_2', angle_h_2)
    # print('angle_v_2', angle_v_2)
    angles = angle_h_1, angle_v_1, angle_h_2, angle_v_2
    n = 0
    for angle in angles:
        # print('*'*40)
        # print(angle.std()/angle.mean())
        if len(angle) > n:
            n = len(angle)
            result = angle.mean()
    # print(result)

    return result

