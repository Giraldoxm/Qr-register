import pyqrcode
import png
from pyqrcode import QRCode

con=1234
while con<= 1235:
    roster=con
    id= '76' + str(con)
    qr=pyqrcode.create(71 and id, error='L')
    qr.png('L' + str(roster)+ '.png', scale=6)
    con=con+1