import RPi.GPIO as GPIO
from imutils.video import VideoStream
from pyzbar import pyzbar
import argparse
import imutils
import time
import cv2
import gspread
from datetime import *
import socket

# The Server Part (Authenticating Google Sheets)
from oauth2client.service_account import ServiceAccountCredentials
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('file_name.json',scope) #Your created json file.
gc = gspread.authorize(credentials)

today = datetime.today().strftime("%d-%m-%Y")

s = []
k=[]
# Opening Google sheet - Attendance
wb = gc.open('Attendance')
worksheets = ((gc.open('Attendance').worksheets()))
list_of_sheets = '{}'.format(worksheets)
print(worksheets,list_of_sheets)
l = list_of_sheets.replace('<','')
l = l.replace('>','')
l = l.replace('[','' )
l = l.replace(']','')
s = (l.split(","))

for i in s:
    l = (i.split("'"))
    k.append(l[1])

if today not in k:
       wb=gc.open('Attendance').add_worksheet(title=today,rows="100",cols="100")
else:
      wb = gc.open('Attendance').worksheet(today)

# Header
wb.update_cell(1, 1, "DATE")
wb.update_cell(1, 2, "ROLL")
wb.update_cell(1, 3, "NAME")
wb.update_cell(1,4,"IN")

t= datetime.now().strftime('%H:%M:%S') #current time
database={'Roll_no':'Name' , '17P211':'JOHNSON'}  # Create your database

GPIO.setmode(GPIO.BCM) #Setting up GPIO pins in RaspberryPi
GPIO.setwarnings(False)

try:
    def sheets(barcode):
        # led READY TO SCAN OFF
        GPIO.setwarnings(False)
        existing_rollno = list(wb.col_values(2))
        roll_no = barcode

        print(roll_no)
        for k, v in database.items():
            if (roll_no == k):
                if (roll_no in existing_rollno):
                    row_no = existing_rollno.index(roll_no) + 1
                    print(row_no)
                    l1 = list(wb.row_values(row_no)) #Elements of the row containing scanned roll_no
                    print(l1)
                    list_index = len(l1)
                    wb.update_cell(row_no, list_index + 1, t)
                    if (list_index % 2 == 0):
                        # led IN
                        wb.update_cell(1, list_index + 1, "OUT")
                    else:
                        # led OUT
                        wb.update_cell(1, list_index + 1, "IN")

                else:
                    # led IN
                    wb.append_row([today, k, v, t])
        else:
            print("Invalid Roll_no")

        time.sleep(5)
        scan()


    def scan():
        # led READY TO SCAN ON
        GPIO.setwarnings(False)
        vs = VideoStream(usePiCamera=True).start()
        time.sleep(2.0)
        # loop over the frames from the video stream
        while True:
            # grab the frame from the threaded video stream and resize it to
            # have a maximum width of 400 pixels
            frame = vs.read()
            frame = imutils.resize(frame, width=400)

            # find the barcodes in the frame and decode each of the barcodes
            barcodes = pyzbar.decode(frame)

            # loop over the detected barcodes
            for barcode in barcodes:
                # extract the bounding box location of the barcode and draw
                # the bounding box surrounding the barcode on the image
                (x, y, w, h) = barcode.rect
                cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 0, 255), 2)

                # the barcode data is a bytes object so if we want to draw it
                # on our output image we need to convert it to a string first
                barcodeData = barcode.data.decode("utf-8")
                barcodeType = barcode.type

        vs.stop()
        sheets(barcodeData)


    def internet_connection():
        try:
            # connect to the host -- tells us if the host is actually
            # reachable
            socket.create_connection(("www.google.com", 80))
            return True
        except OSError:
            pass
        return False

    if (internet_connection() == True):
        scan()
    else:
        # led Indication
        print("No internet connection")
except:
    # led Indication
    print("Temporarily Out of Service")
