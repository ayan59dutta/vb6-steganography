import numpy as np
from PIL import Image

def decode(password, image_path):
    image = Image.open(image_path)
    img = np.array(image)
    code_text_hex = ''
    flag_pass = True
    eof, flag_eof = '', False
    sof, flag_sof = '', False
    for r, g, b in img[0][:2]:
        sof += (hex(r)[-1] + hex(g)[-1] + hex(b)[-1])
    if int(sof, 16) == 1114111:
        flag_sof = True
    if flag_sof == True:
        i = 0
        pass_hex = ''
        for character in password:
            pass_hex += (hex(ord(character)))[2:]
        for pix_row in img:
            for r, g, b in pix_row:
                code_text_hex += (hex(r)[-1] + hex(g)[-1] + hex(b)[-1])
                if i == 0 and code_text_hex.count(hex(1114111)[2:]) == 2:
                    if code_text_hex[6:code_text_hex.index(hex(1114111)[2:], 6)] != pass_hex:
                        flag_pass = False
                    i += 1
                if i > 0 and code_text_hex.count(hex(1114111)[2:]) == 3:
                    flag_eof = True
                    break
            if flag_eof == True:
                break
        if flag_eof == True and flag_pass == True:
            text = ''
            text_hex = code_text_hex[code_text_hex.index(hex(1114111)[2:], 6)+6:-6]
            len_text_hex = len(text_hex)
            i = 0
            while i+1 < len_text_hex:
                text += chr(int(text_hex[i:i+2], 16))
                i += 2
            if ord(text[-1]) == 16:
                text = text[:-1]
            return (True, text)
        elif flag_eof == False and flag_pass == True:
            error_msg = 'Cannot extract text. End not reached.'
            return (False, error_msg)
        else:
            error_msg = 'Wrong Password.'
            return (False, error_msg)
    else:
        error_msg = 'Image contains no hidden text.'
        return (False, error_msg)

def decrypt(ciphertext, key):
    plaintext = ''
    for character in ciphertext:
        plaintext += chr((ord(character)-key)%256)
    return plaintext

if __name__ == '__main__':
    fp = open('ipd.txt')
    image_path = fp.readline()[:-1]
    password = fp.readline()[:-1]
    fp.close()
    flag_success = decode(password, decrypt(image_path, 0))
    fp = open('opd.txt', 'w')
    fp.write(str(flag_success[0])+'\n')
    fp.write(flag_success[1])
    fp.close()
