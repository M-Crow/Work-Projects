# barcode scanner 
from cmath import rect
import cv2
import numpy as np 
from pyzbar.pyzbar import decode

# capture the video from device camera
cap = cv2.VideoCapture(0)
while True:
    ret, frame = cap.read()
    cv2.imshow('image', frame)
    code = cv2.waitkey(10)
    if code == ord('q'):
        break

# decorder function - decodes barcode and QR code from a given image
def decoder(image):
    gray_img = cv2.cvtColor(image,0)
    barcode = decode(gray_img)

    for obj in barcode:
        points = obj.polygon
        (x,y,w,h) = obj.rect
        pts = np.array(points, np.int32)
        pts = pts.reshare((-1, 1, 2))
        cv2.polylines(image, [pts], True, (0, 255, 0), 3)

        barcodeData = obj.data.decode("utf-8")
        barcodeType = obj.type
        string = "Data: " + str(barcodeData) + " | Type: " + str(barcodeType)

        cv2.putText(frame, string, (x,y), cv2.FONT_HERSHEY_SIMPLEX,0.8,(0,0,255),2)
        print("Barcode: "+barcodeData +" | Type: "+barcodeType)