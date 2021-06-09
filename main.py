from PIL import Image,ImageDraw, ImageFont
from PyPDF2 import PdfFileMerger
import hashlib
import pyqrcode 
import png 
import openpyxl
from pyqrcode import QRCode 
def qr_code_png(shahex):
    url = pyqrcode.create(shahex)
    return url

def certificate_to_hashcode(certificate):
    hash = hashlib.sha256(certificate).hexdigest()
    return hash

def insert_participant_data(id,name,namex,namey,attfrom,attfromx,attfromy,template,namefontsize=100,attfromfontsize=75):
    im = Image.open(template)
    certicate_template = ImageDraw.Draw(im)
    name_location = (namex,namey)
    attfrom_location = (attfromx,attfromy)
    text_color = (0, 0, 0)
    name_font = ImageFont.truetype("Palatino Linotype.ttf", namefontsize)
    attfrom_font = ImageFont.truetype("Palatino Linotype.ttf", attfromfontsize)
    certicate_template.text(name_location,name, fill = text_color, font = name_font)
    certicate_template.text(attfrom_location,attfrom, fill = text_color, font = attfrom_font)

    QRCode_Data = "ID: "+str(id)+"\nName: "+str(name)+"\nFrom: "+str(attfrom)

    url = qr_code_png(QRCode_Data)
    url.png('code.png', scale = 8) 
    qr_code = Image.open("code.png")
    qr_location = (1000,2800)
    im.paste(qr_code,qr_location)
    sha256_code = certificate_to_hashcode(im.tobytes())
    im.save("gen/"+id+"_"+name+"_"+attfrom+".pdf")

template = "template.jpg"

wb_obj = openpyxl.load_workbook("data.xlsx")
sheet = wb_obj.active
for row in sheet.iter_rows():
    namefontsize=80
    attfromfontsize=75
    id = row[1].value
    name = row[2].value
    att_from=row[3].value
    if(len(att_from)>25):
        attfromfontsize=60
    if(len(att_from)>35):
        attfromfontsize=50
    if(len(name)>13):
        namefontsize=60
    insert_participant_data(row[1].value,row[2].value,1000,1950,row[3].value,210,2100,template,namefontsize,attfromfontsize)