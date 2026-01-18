import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser
from tkinterdnd2 import DND_FILES, TkinterDnD
import qrcode
from PIL import Image, ImageTk
from pyzbar.pyzbar import decode
import cv2
import os
from datetime import datetime
from reportlab.pdfgen import canvas as pdf_canvas
import numpy as np
import threading
import time
import pyttsx3

class QRCodeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("QR Code App")
        self.root.geometry("640x740")
        self.dark_mode = True
        self.bg_image = None
        self.fg_color = "black"  # Foreground color for QR
        self.bg_color = "white"  # Background color for QR
        self.border_color = "black"  # Default border color
        self.border_size = 10  # Default border size
        self.logo_path = None
        self.qr_image = None

        self.load_background()
        self.setup_ui()

        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind("<<Drop>>", self.handle_drop)

    def load_background(self):
        if os.path.exists("assets/background.png"):
            bg = Image.open("assets/background.png").resize((640, 740))
            self.bg_image = ImageTk.PhotoImage(bg)

    def setup_ui(self):
        if self.bg_image:
            self.bg_label = tk.Label(self.root, image=self.bg_image)
            self.bg_label.place(x=0, y=0, relwidth=1, relheight=1)
        else:
            self.bg_label = tk.Label(self.root)

        font_style = ("Montserrat", 12, "bold")

        self.entry = tk.Entry(self.root, font=font_style, bg="#f4f4f4", fg="#333", relief="flat")
        self.entry.place(x=70, y=20, width=500, height=30)

        self.buttons = []
        btn_data = [
            ("Generate QR", self.generate_qr),
            ("Choose Logo (PNG)", self.choose_logo),
            ("Change QR Color", self.choose_color),
            ("Choose Image to Scan", self.choose_image_to_scan),
            ("Scan via Webcam", self.scan_via_webcam),
            ("Toggle Theme", self.toggle_theme),
            ("Save as PNG", self.save_qr_png),
            ("Save as JPG", self.save_qr_jpg),
            ("Save as PDF", self.save_qr_pdf),
            ("Speak QR Data", self.speak_qr_data)
        ]

        y_base = 60
        for i, (text, command) in enumerate(btn_data):
            btn = tk.Button(
                self.root, text=text, command=lambda c=command: self.animated_click(c),
                font=("Montserrat", 10), bg="#4CAF50", fg="black", activebackground="#45a049", relief="flat",
                cursor="hand2"
            )
            btn.place(x=70 + (i % 2) * 280, y=y_base + 40 * (i // 2), width=240, height=30)
            self.buttons.append(btn)

        self.canvas = tk.Canvas(self.root, width=320, height=320, bg="white", bd=0, highlightthickness=2, highlightbackground="#ccc")
        self.canvas.place(x=160, y=400)

    def animated_click(self, command):
        self.root.config(cursor="watch")
        self.root.update()
        time.sleep(0.1)
        command()
        self.root.config(cursor="")

    def toggle_theme(self):
        self.dark_mode = not self.dark_mode
        bg_color = "#1e1e1e" if self.dark_mode else "#ffffff"
        self.root.configure(bg=bg_color)
        self.bg_label.configure(bg=bg_color)
        self.entry.configure(bg="#f4f4f4" if not self.dark_mode else "#333", fg="#000")
        for btn in self.buttons:
            btn.configure(bg="#4CAF50", fg="black", activebackground="#45a049")
        self.canvas.configure(bg="white")

    def generate_qr(self):
        # Start QR generation in a new thread
        thread = threading.Thread(target=self._generate_qr_thread)
        thread.daemon = True
        thread.start()

    def _generate_qr_thread(self):
        data = self.entry.get()
        if not data:
            messagebox.showwarning("Input Error", "Please enter some text.")
            return

        qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_H)
        qr.add_data(data)
        qr.make(fit=True)
        img = qr.make_image(fill_color=self.fg_color, back_color=self.bg_color).convert('RGB')

        # Add logo if set
        if self.logo_path:
            logo = Image.open(self.logo_path).convert("RGBA")
            logo.thumbnail((80, 80))  # Resize logo to fit inside QR code
            pos = ((img.size[0] - logo.size[0]) // 2, (img.size[1] - logo.size[1]) // 2)
            img.paste(logo, pos, mask=logo)  # Use logo transparency mask

        # Add border
        img_with_border = self.add_border(img)

        self.qr_image = img_with_border
        self.root.after(0, self.display_qr, img_with_border)

    def display_qr(self, img):
        img = img.resize((320, 320), Image.LANCZOS)
        self.tk_img = ImageTk.PhotoImage(img)
        self.canvas.delete("all")
        self.canvas.create_image(160, 160, image=self.tk_img)

    def add_border(self, img):
        new_img_size = (img.size[0] + 2 * self.border_size, img.size[1] + 2 * self.border_size)
        new_img = Image.new("RGB", new_img_size, self.border_color)
        new_img.paste(img, (self.border_size, self.border_size))
        return new_img

    def choose_logo(self):
        path = filedialog.askopenfilename(filetypes=[("PNG Images", "*.png")])
        if path:
            self.logo_path = path

    def choose_color(self):
        color = colorchooser.askcolor()[1]
        if color:
            self.fg_color = color

    def save_qr_png(self):
        self.save_qr("PNG", ".png")

    def save_qr_jpg(self):
        self.save_qr("JPEG", ".jpg")

    def save_qr_pdf(self):
        if self.qr_image:
            path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if path:
                img_path = path.replace(".pdf", ".png")
                self.qr_image.save(img_path)
                c = pdf_canvas.Canvas(path)
                c.drawImage(img_path, 100, 500, width=300, height=300)
                c.save()
                os.remove(img_path)
        else:
            messagebox.showerror("No QR", "Generate a QR code first.")

    def save_qr(self, format, ext):
        if self.qr_image:
            path = filedialog.asksaveasfilename(defaultextension=ext, filetypes=[(f"{format} files", f"*{ext}")])
            if path:
                self.qr_image.save(path, format=format)
        else:
            messagebox.showerror("No QR", "Generate a QR code first.")

    def choose_image_to_scan(self):
        path = filedialog.askopenfilename(filetypes=[("Image Files", "*.png *.jpg *.jpeg")])
        if path:
            self.decode_qr(path)

    def decode_qr(self, path):
        try:
            img = Image.open(path)
            result = decode(img)
            if result:
                data = result[0].data.decode('utf-8')
                messagebox.showinfo("QR Code Data", data)
            else:
                messagebox.showwarning("Not Found", "No QR code found in image.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def handle_drop(self, event):
        path = event.data.strip('{}')
        if os.path.isfile(path):
            self.decode_qr(path)

    def scan_via_webcam(self):
        cap = cv2.VideoCapture(0)
        while True:
            ret, frame = cap.read()
            if not ret:
                break
            decoded_objs = decode(frame)
            for obj in decoded_objs:
                points = obj.polygon
                if len(points) > 4:
                    hull = cv2.convexHull(np.array([point for point in points], dtype=np.float32))
                    hull = list(map(tuple, np.squeeze(hull)))
                else:
                    hull = points
                n = len(hull)
                for j in range(0, n):
                    cv2.line(frame, hull[j], hull[(j + 1) % n], (255, 0, 0), 3)
                data = obj.data.decode("utf-8")
                cv2.putText(frame, data, (20, 50), cv2.FONT_HERSHEY_SIMPLEX, 1, (0,255,0), 2)
                messagebox.showinfo("Webcam Scan Result", data)
                self.announce(data)  # Announce the result
                cap.release()
                cv2.destroyAllWindows()
                return
            cv2.imshow("QR Scanner", frame)
            if cv2.waitKey(1) & 0xFF == ord('q'):
                break
        cap.release()
        cv2.destroyAllWindows()

    def announce(self, data):
        engine = pyttsx3.init()
        engine.say(f"QR Code Data: {data}")
        engine.runAndWait()

    def speak_qr_data(self):
        if self.qr_image:
            data = self.entry.get()
            engine = pyttsx3.init()
            engine.say(data)
            engine.runAndWait()
        else:
            messagebox.showwarning("No QR", "Generate a QR code first.")

if __name__ == '__main__':
    app = TkinterDnD.Tk()
    QRCodeApp(app)
    app.mainloop()
