import cv2
import sqlite3
import pandas as pd
import qrcode
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import messagebox
from datetime import datetime
from pathlib import Path
from tkinter import Tk, Canvas, Entry, Button, PhotoImage, Label
from PIL import Image, ImageTk


OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = OUTPUT_PATH / Path(r"C:\Users\night\OneDrive\Desktop\CPE Project\frame0")

DB_FILE = "attendance.db"
EXCEL_FILE = "attendance.xlsx"

def initialize_database():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE,
            timestamp TEXT
        )
    ''')
    conn.commit()
    conn.close()

def mark_attendance(data):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT * FROM attendance WHERE name = ?", (data,))
    if c.fetchone():
        conn.close()
        return False 
    c.execute("INSERT INTO attendance (name, timestamp) VALUES (?, ?)", (data, timestamp))
    conn.commit()
    conn.close()
    return True

def export_to_excel():
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT name, timestamp FROM attendance", conn)
    df.to_excel(EXCEL_FILE, index=False)
    conn.close()
    print(f"Attendance records exported to {EXCEL_FILE}")

def generate_qr_code(name, student_id):
    data = f"{name} ({student_id})"
    qr = qrcode.QRCode(
        version=1,
        box_size=10,
        border=4
    )
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill='black', back_color='white')
    filename = f"{name}_{student_id}_qr.png"
    img.save(filename)
    return filename


def relative_to_assets(path: str) -> Path:
    return ASSETS_PATH / Path(path)

class QRGeneratorWindow(tk.Toplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master

        self.title("QR Code Generator")
        self.geometry("918x598")
        self.configure(bg="#860101")
        self.resizable(False, False)
        
        self.canvas = Canvas(
            self,
            bg="#860101",
            height=598,
            width=918,
            bd=0,
            highlightthickness=0,
            relief="ridge"
        )


        self.canvas.place(x=0, y=0)

     
        self.canvas.create_text(59.0, 164.0, anchor="nw", text="Name:",
                                fill="#FFFFFF", font=("Limelight Regular", 36 * -1))
        self.canvas.create_text(59.0, 283.0, anchor="nw", text="Student ID:",
                                fill="#FFFFFF", font=("Roboto Regular", 36 * -1))
        self.canvas.create_text(99.0, 21.0, anchor="nw", text="Attendance Monitoring",
                                fill="#FFFFFF", font=("RobotoCondensed Bold", 64 * -1))


        self.name_entry_image = PhotoImage(file=relative_to_assets("entry_1.png"))
        self.canvas.create_image(313.0, 236.0, image=self.name_entry_image)

        self.id_entry_image = PhotoImage(file=relative_to_assets("entry_2.png"))
        self.canvas.create_image(313.0, 357.0, image=self.id_entry_image)


        self.name_entry = Entry(self, bd=0, bg="#D9D9D9", fg="#000716", highlightthickness=0)
        self.name_entry.place(x=69.0, y=206.0, width=488.0, height=58.0)

        self.id_entry = Entry(self, bd=0, bg="#D9D9D9", fg="#000716", highlightthickness=0)
        self.id_entry.place(x=69.0, y=327.0, width=488.0, height=58.0)


        self.qr_label = Label(self, bg="#860101")
        self.qr_label.place(x=700, y=200, width=150, height=150) 


        self.button_image = PhotoImage(file=relative_to_assets("asd.png"))
        self.button = Button(self, image=self.button_image, borderwidth=0, highlightthickness=0,
                             command=self.generate_qr, relief="flat")
        self.button.place(x=287.0, y=435.0, width=352.0, height=82)


        self.button_image_back = PhotoImage(file=relative_to_assets("button_1.png"))
        self.button_back = Button(self, image=self.button_image_back, borderwidth=0,
                                  highlightthickness=0, command=self.on_back, relief="raised")
        self.button_back.place(x=32.0, y=33.0, width=51.0, height=51.0)

    def generate_qr(self):
        name = self.name_entry.get().strip()
        student_id = self.id_entry.get().strip()
        if not name or not student_id:
            messagebox.showerror("Input Error", "Please enter both name and ID.")
            return

        filename = generate_qr_code(name, student_id)
        img = Image.open(filename)
        img = img.resize((150, 150))
        photo = ImageTk.PhotoImage(img)
        
        self.qr_label.config(image=photo)
        self.qr_label.image = photo

        messagebox.showinfo("Success", f"QR Code generated and saved as {filename}")
    
    def on_back(self):
        """Restores MainApplication and closes QRGeneratorWindow."""
        self.master.deiconify() 
        self.destroy() 



class AttendanceScannerWindow(tk.Toplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master

        self.title("QR Attendance Scanner")
        self.geometry("918x598")
        self.configure(bg="#860101")
        self.resizable(False, False)

        self.canvas = Canvas(
            self,
            bg="#860101",
            height=598,
            width=918,
            bd=0,
            highlightthickness=0,
            relief="ridge"
        )
        self.canvas.place(x=0, y=0)

        self.canvas.create_text(
            95.0, 20.0, anchor="nw",
            text="Attendance Monitoring",
            fill="#FFFFFF",
            font=("RobotoCondensed Bold", 64 * -1)
        )

        self.video_label = Label(self)
        self.video_label.place(x=50.0, y=175.0, width=400, height=300)

        self.button_image_back = PhotoImage(file=relative_to_assets("button_1.png"))
        self.button_back = Button(self, image=self.button_image_back, borderwidth=0,
                                  highlightthickness=0, command=self.on_back, relief="raised")
        self.button_back.place(x=32.0, y=33.0, width=51.0, height=51.0)


        self.detector = cv2.QRCodeDetector()

 
        self.start_video()


        self.protocol("WM_DELETE_WINDOW", self.on_back)

    def start_video(self):
        self.cap = cv2.VideoCapture(0)
        if not self.cap.isOpened():
            messagebox.showerror("Camera Error", "Could not open camera.")
            return

        # Set camera resolution (optional)
        self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
        self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)

        self.update_frame()

    def update_frame(self):
        ret, frame = self.cap.read()
        if ret and frame is not None:
            data, bbox, _ = self.detector.detectAndDecode(frame)
            if bbox is not None and data:
                if mark_attendance(data):
                    print(f"Attendance marked for: {data}")
                else:
                    print(f"Already marked attendance for: {data}")

                n = len(bbox)
                for i in range(n):
                    pt1 = tuple(map(int, bbox[i][0]))
                    pt2 = tuple(map(int, bbox[(i + 1) % n][0]))
                    cv2.line(frame, pt1, pt2, (0, 255, 0), 3)

                cv2.putText(frame, data, (int(bbox[0][0][0]), int(bbox[0][0][1]) - 10),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.9, (0, 255, 0), 2)

            cv2image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGBA)
            img = Image.fromarray(cv2image)
            imgtk = ImageTk.PhotoImage(image=img)

            self.video_label.imgtk = imgtk  # keep reference
            self.video_label.configure(image=imgtk)
        else:
            print("Failed to grab frame from camera.")

        self.after(10, self.update_frame)

    def on_back(self):
        if hasattr(self, 'cap') and self.cap.isOpened():
            self.cap.release()
        self.master.deiconify()
        export_to_excel()
        self.destroy()



class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("QR Attendance System")
        self.geometry("918x598")
        self.configure(bg="#860101")
    
        canvas = Canvas(
            self,
            bg = "#860101",
            height = 598,
            width = 918,
            bd = 0,
            highlightthickness = 0,
            relief = "ridge"
        )

        canvas.place(x = 0, y = 0)

        self.button_image_1 = PhotoImage(file=relative_to_assets("asd.png"))
        self.button_1 = Button(
            image=self.button_image_1,
            borderwidth=0,
            highlightthickness=0,
            command=self.open_qr_generator,
            relief="flat"
        )
        self.button_1.place(
            x=264.0,
            y=234.0,
            width=390.0,
            height=91.0
        )

        self.button_image_2 = PhotoImage(file=relative_to_assets("start.png"))
        self.button_2 = Button(
            image=self.button_image_2,
            borderwidth=0,
            highlightthickness=0,
            command=self.open_attendance_scanner,
            relief="flat"
)
        self.button_2.place(
            x=264.0,
            y=371.0,
            width=390.0,
            height=91.0
        )

        image_image_1 = PhotoImage(
            file=relative_to_assets("image_1.png"))
        image_1 = canvas.create_image(
            54.0,
            59.0,
            image=image_image_1
        )

        canvas.create_text(
            99.0,
            19.0,
            anchor="nw",
            text="Attendance Monitoring",
            fill="#FFFFFF",
            font=("Inter Bold", 64 * -1)
        )
    
    def open_qr_generator(self):

        self.withdraw()  # Hide main window
        qr_window = QRGeneratorWindow(self)  
        qr_window.protocol("WM_DELETE_WINDOW", lambda: self.on_qr_close())  # Handle when closed

    def on_qr_close(self):

        self.deiconify()  # Show main window again


    def open_attendance_scanner(self):

        self.withdraw()  # Hide main window
        qr_window = AttendanceScannerWindow(self)  
        qr_window.protocol("WM_DELETE_WINDOW", lambda: self.on_qr_close())  # Handle when closed


if __name__ == "__main__":
    #initialize_database()
    app = MainApplication()
    app.mainloop()