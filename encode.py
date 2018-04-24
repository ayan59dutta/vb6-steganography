import numpy as np
from PIL import Image
import os

def encode(password, text, image_src_path, image_dest_path):
    image = Image.open(image_src_path)
    img = np.array(image)
    code_text = (chr(1114111)) + password + (chr(1114111)) + text + (chr(1114111))
    code_text_hex = ''
    for character in code_text:
        asc = ord(character)
        if asc < 16:
            code_text_hex += ('0' + hex(asc)[2:])
        else:
            code_text_hex += (hex(asc))[2:]
    len_hex = len(code_text_hex)
    i = 0
    for pix_i in range(len(img)):
        for pix_j in range(len(img[pix_i])):
            r, g, b = img[pix_i][pix_j][0], img[pix_i][pix_j][1], img[pix_i][pix_j][2]
            if i < len_hex:
                r = int(hex(r)[:-1] + code_text_hex[i], 16)
                img[pix_i][pix_j][0] = r
                i += 1
            else:
                break
            if i < len_hex:
                g = int(hex(g)[:-1] + code_text_hex[i], 16)
                img[pix_i][pix_j][1] = g
                i += 1
            else:
                break
            if i < len_hex:
                b = int(hex(b)[:-1] + code_text_hex[i], 16)
                img[pix_i][pix_j][2] = b
                i += 1
            else:
                break
    if i == len_hex:
        img = Image.fromarray(img)
        img.save(image_dest_path)
        flag_success = (True, 'Text successfully hidden within the image.')
    else:
        error_msg = 'Text too large for the given image.'
        flag_success = (False, error_msg)
    return flag_success

if __name__ == '__main__':
    try:
        fp = open('ipe.txt')
        image_src_path = fp.readline()[:-1]
        image_dest_path = fp.readline()[:-1]
        password = fp.readline()[:-1]
        text = ''
        line = fp.readline()
        while line != '':
            text += line[:-1]
            line = fp.readline()
        fp.close()
        os.remove('ipe.txt')
        flag_success = encode(password, text, image_src_path, image_dest_path)
        fp = open('ope.txt', 'w')
        fp.write(str(flag_success[0]) + '\n')
        fp.write(flag_success[1] + '\n')
        fp.close()
    except:
        fp = open('ope.txt', 'w')
        fp.write(str(False) + '\n')
        fp.write('Encoding not possible in this image.' + '\n')
        fp.close()
