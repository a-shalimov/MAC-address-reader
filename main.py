from tkinter import *
from openpyxl import Workbook
import cv2
from PIL import Image as PILImage, ImageTk
import pytesseract
from pyzbar import pyzbar
import re

root = Tk()
root.title("MAC Reader")

mac_list_frame = LabelFrame(root, text="MAC List")
mac_list_frame.pack(side=LEFT, fill=Y)

mac_listbox = Listbox(mac_list_frame)
mac_listbox.pack(fill=BOTH, expand=True)

video_frame = LabelFrame(root, text="Video Feed")
video_frame.pack(side=LEFT)

mac_label = Label(video_frame, text="MAC: ", font=("Arial", 24))
mac_label.pack()

video_label = Label(video_frame)
video_label.pack()

cap = cv2.VideoCapture(0)

wb = Workbook()
ws = wb.active

def get_frame():
    """
    Capture a frame from the video feed and display it in the video label.
    """
    ret, frame = cap.read()
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    video_label.imgtk = ImageTk.PhotoImage(image=PILImage.fromarray(gray))
    video_label.configure(image=video_label.imgtk)
    root.after(10, get_frame)


def capture():
    """
    Capture a frame from the video feed and try to read a QR code or perform OCR to recognize a MAC address.
    """
    ret, frame = cap.read()
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    codes = pyzbar.decode(gray)
    if codes:
        mac_address = codes[0].data.decode()
        mac_label.configure(text=f"MAC: {mac_address}")
        return

    text = pytesseract.image_to_string(gray)
    
    mac_address_search = re.search(r'(?:MAC(?:ID)?|MAC\s*:?\s*)([0-9A-Fa-f]{12})', text)
    
    if mac_address_search:
        mac_address = mac_address_search.group(1)
        mac_label.configure(text=f"MAC: {mac_address}")
        return


def approve():
    """
    Approve the recognized MAC address and add it to the listbox and worksheet.
    """
    mac_address = mac_label.cget("text").split()[-1]

    if mac_address not in mac_listbox.get(0, END):
        mac_listbox.insert(END, mac_address)
        ws.append([mac_address])

    mac_label.configure(text="MAC: ")


def key_press(event):
    """
    Handle key press events to capture, approve or deny a MAC address.
    
    Parameters: 
        event (obj): The key press event.
    """
    if event.char == 'c':
        capture()
    elif event.char == 'a':
        approve()

root.bind("<Key>", key_press)

button_width = int(mac_listbox.cget("width"))

capture_button = Button(root, text="Capture", command=capture, width=button_width)
capture_button.pack(side=TOP)

approve_button = Button(root, text="Approve", command=approve, width=button_width)
approve_button.pack(side=TOP)


get_frame()

root.mainloop()

wb.save("mac_addresses.xlsx")
