from libraries import *
def get_qr(link_qr,qr_name,color_select):

    if color_select == 0 :
        logo = Image.open('QR/img/logo.jpg')
    else:
        logo = Image.open('QR/img/logo_black.jpg')

    basewidth = 600
    wp = (basewidth/float(logo.size[0]))
    hs = int((float(logo.size[1])*float(wp)))
    logo = logo.resize((basewidth,hs), Image.Resampling.LANCZOS)
    qr_big = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_H,box_size = 60)
    qr_big.add_data(link_qr)
    qr_big.make()

    if color_select == 0:
        img_qr_big = qr_big.make_image(fill_color = '#00B2B5' , back_color="white").convert('RGB')
    else:
        img_qr_big = qr_big.make_image(fill_color = '#000000' , back_color="white").convert('RGB')

    pos = ((img_qr_big.size[0]- logo.size[0]) // 2, (img_qr_big.size[1] - logo.size[1]) // 2)
    img_qr_big.paste(logo, pos)

    if color_select == 0:
        img_qr_big.save('QR/qr_code/%s_qr.jpg'%(qr_name))
    else:
        img_qr_big.save('QR/qr_code/%s_qr_black.jpg'%(qr_name))