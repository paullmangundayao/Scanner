import tkinter as tk
from tkinter import filedialog
import cv2
from PIL import Image, ImageTk
import numpy as np
from datetime import datetime
import os
import pyttsx3
import threading
from fpdf import FPDF
from imutils.perspective import four_point_transform
import queue

class ScannerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("A.F.P.M.B.A.I Scanner Application")
        
        # Initialize scanner variables
        self.WIDTH = 1920
        self.HEIGHT = 1080
        self.A4_width = 210
        self.A4_height = 297
        self.scanning = False
        self.selecting = False
        self.drawing = False
        self.pdf = None
        self.whitened_rects = []
        self.preview = None
        self.scanned = None
        self.rect_start = None
        self.rect_end = None
        
        # Initialize text-to-speech
        self.engine = pyttsx3.init()
        self.engine.setProperty('rate', 175)
        
        # Create "Scanned Files" directory in Documents
        self.destination_folder = os.path.join(os.path.expanduser('~'), 'Downloads')
        os.makedirs(self.destination_folder, exist_ok=True)

        # Initialize camera index
        self.camera_index = 1 
        
        # Setup camera
        self.cap = cv2.VideoCapture(self.camera_index, cv2.CAP_DSHOW)
        self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, self.WIDTH)
        self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, self.HEIGHT)
        
        # Initialize speech queue and thread
        self.speech_queue = queue.Queue()
        self.speech_thread = threading.Thread(target=self.process_speech_queue)
        self.speech_thread.daemon = True  # Allow thread to exit when the main program exits
        self.speech_thread.start()

        # Create GUI layout
        self.create_gui_elements()
        
        # Start video stream
        self.update_video()
        
        # Initial speak
        self.speak("Initializing A.F.P.M.B.A.I. Scanner Application")

    def process_speech_queue(self):
        while True:
            text = self.speech_queue.get()
            if text is None:  # Exit signal
                break
            self.engine.say(text)
            self.engine.runAndWait()

    def speak(self, text):
        self.speech_queue.put(text)

    def create_gui_elements(self):
        # Create main container with gray background
        self.root.configure(bg='#f0f0f0')
        
        # Create main frames
        self.top_frame = tk.Frame(self.root, bg='#f0f0f0')
        self.top_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.bottom_frame = tk.Frame(self.root, bg='#f0f0f0')
        self.bottom_frame.pack(fill=tk.X, padx=10, pady=10)

        # Status text at the top of the top frame
        self.status_var = tk.StringVar(value="Starting a new scan session")
        self.status_label = tk.Label(self.top_frame, 
                                   textvariable=self.status_var,
                                   bg='#4ae816',  # Green background
                                   fg='#000000',  # Blue text color
                                   font=('Arial', 14, 'bold'),  # Larger and bold font
                                   relief=tk.RAISED,  # Add a raised border
                                   borderwidth=2,  # Border width
                                   padx=10,  # Padding for better spacing
                                   pady=10)
        self.status_label.pack(pady=(0, 10))

        # Create a frame for the side-by-side previews
        self.preview_container = tk.Frame(self.top_frame, bg='#f0f0f0')
        self.preview_container.pack(fill=tk.BOTH, expand=True)

        # Camera preview frame (left half of the screen)
        self.camera_frame = tk.Frame(self.preview_container, bg='black')
        self.camera_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.camera_label = tk.Label(self.camera_frame, bg='black')
        self.camera_label.pack(fill=tk.BOTH, expand=True)

        # Document preview frame (right half of the screen)
        self.document_frame = tk.Frame(self.preview_container, bg='#e8e8e8')
        self.document_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.document_label = tk.Label(self.document_frame, 
                                     text="Document Preview",
                                     bg='#e8e8e8')
        self.document_label.place(relx=0.5, rely=0.5, anchor='center')

        # Buttons
        self.create_buttons()

        # Bind window resize event
        self.root.bind("<Configure>", self.on_window_resize)

    def on_window_resize(self, event):
        # Update the previews when the window is resized
        self.update_video()

    def create_buttons(self):
        button_frame = tk.Frame(self.bottom_frame, bg='#f0f0f0')
        button_frame.pack(fill=tk.X, pady=10)

        # Button styles
        scan_button = tk.Button(button_frame,
                              text="Add page",
                              command=self.handle_scan,
                              bg='#4285f4',
                              fg='white',
                              font=('Arial', 11),
                              width=20,
                              height=2)
        scan_button.pack(side=tk.LEFT, padx=5, pady=5)

        save_button = tk.Button(button_frame,
                              text="Save PDF",
                              command=self.handle_save,
                              bg='#34a853',
                              fg='white',
                              font=('Arial', 11),
                              width=20,
                              height=2)
        save_button.pack(side=tk.LEFT, padx=5, pady=5)

        edit_button = tk.Button(button_frame,
                              text="Edit",
                              command=self.handle_edit,
                              bg='#e8e8e8',
                              font=('Arial', 11),
                              width=20,
                              height=2)
        edit_button.pack(side=tk.LEFT, padx=5, pady=5)

        folder_button = tk.Button(button_frame,
                                text="Open Folder",
                                command=self.handle_open_folder,
                                bg='#e8e8e8',
                                font=('Arial', 11),
                                width=20,
                                height=2)
        folder_button.pack(side=tk.LEFT, padx=5, pady=5)

        # New button to select preferred folder
        select_folder_button = tk.Button(button_frame,
                                        text="Select Preferred Folder",
                                        command=self.handle_select_folder,
                                        bg='#e8e8e8',
                                        font=('Arial', 11),
                                        width=20,
                                        height=2)
        select_folder_button.pack(side=tk.LEFT, padx=5, pady=5)

        # Add a camera switch button
        camera_switch_button = tk.Button(button_frame,
                                       text="Switch Camera",
                                       command=self.handle_switch_camera,
                                       bg='#e8e8e8',
                                       font=('Arial', 11),
                                       width=20,
                                       height=2)
        camera_switch_button.pack(side=tk.LEFT, padx=5, pady=5)

        exit_button = tk.Button(button_frame,
                              text="Exit",
                              command=self.handle_exit,
                              bg='#ea4335',
                              fg='white',
                              font=('Arial', 11),
                              width=20,
                              height=2)
        exit_button.pack(side=tk.LEFT, padx=5, pady=5)

    def handle_select_folder(self):
        # Allow user to select a preferred folder
        selected_folder = filedialog.askdirectory()
        if selected_folder:  # Check if the user selected a folder
            self.destination_folder = selected_folder
            self.status_var.set(f"Selected folder: {self.destination_folder}")
            self.speak(f"Selected folder: {self.destination_folder}")

    def handle_switch_camera(self):
        # Release the current camera
        if self.cap.isOpened():
            self.cap.release()

        # Switch to the next camera (assuming you have multiple cameras)
        self.camera_index = (self.camera_index + 1) % 2  # Toggle between 0 and 1
        self.cap = cv2.VideoCapture(self.camera_index, cv2.CAP_DSHOW)
        self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, self.WIDTH)
        self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, self.HEIGHT)

        # Update the status
        self.status_var.set(f"Switched to camera {self.camera_index}")
        self.speak(f"Switched to camera {self.camera_index}")

    def resize_image_to_fit(self, image, max_width, max_height):
        # Resize the image to fit within the given dimensions while maintaining aspect ratio
        height, width = image.shape[:2]
        aspect_ratio = width / height

        if width > max_width or height > max_height:
            if aspect_ratio > 1:  # Landscape
                new_width = max_width
                new_height = int(new_width / aspect_ratio)
            else:  # Portrait
                new_height = max_height
                new_width = int(new_height * aspect_ratio)

            # Ensure new dimensions are valid
            if new_width <= 0 or new_height <= 0:
                new_width = width
                new_height = height

            if new_width > max_width or new_height > max_height:
                if new_width > max_width:
                    new_width = max_width
                    new_height = int(new_width / aspect_ratio)
                if new_height > max_height:
                    new_height = max_height
                    new_width = int(new_height * aspect_ratio)

            # Ensure new dimensions are valid
            if new_width <= 0 or new_height <= 0:
                new_width = width
                new_height = height

            image = cv2.resize(image, (new_width, new_height))
        return image

    def update_video(self):
        if self.cap.isOpened():
            ret, frame = self.cap.read()
            if ret:
                cv2.normalize(frame, frame, 0, 255, cv2.NORM_MINMAX)
                # Frame processing
                width = 1920
                height = 1080
                roi_width = 1300
                roi_height = 1080
                x1 = (width - roi_width) // 2
                y1 = (height - roi_height) // 2
                x2 = x1 + roi_width
                y2 = y1 + roi_height
                
                frame_roi = frame[y1:y2, x1:x2]
                cv2.rectangle(frame, (x1, y1), (x2, y2), (255, 0, 0), 2)

                # Process ROI for document detection
                gray = cv2.cvtColor(frame_roi, cv2.COLOR_BGR2GRAY)
                blur = cv2.GaussianBlur(gray, (5, 5), 0)
                _, threshold = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                
                contours, _ = cv2.findContours(threshold, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
                contours = sorted(contours, key=cv2.contourArea, reverse=True)

                # Find document contour
                for contour in contours:
                    area = cv2.contourArea(contour)
                    if area > 1000:
                        peri = cv2.arcLength(contour, True)
                        approx = cv2.approxPolyDP(contour, 0.015 * peri, True)
                        if len(approx) == 4:
                            adjusted_contour = approx + np.array([x1, y1])
                            cv2.drawContours(frame, [adjusted_contour], -1, (0, 255, 0), 3)
                            
                            # Store the processed image for scanning
                            warped = four_point_transform(frame, adjusted_contour.reshape(4, 2))
                            kernel = np.array([[0, -1, 0], [-1, 5, -1], [0, -1, 0]])
                            sharpened_image = cv2.filter2D(warped, -1, kernel)
                            self.scanned = cv2.cvtColor(sharpened_image, cv2.COLOR_BGR2GRAY)
                            _, self.scanned = cv2.threshold(self.scanned, 128, 255, cv2.THRESH_BINARY)
                            
                            # Create initial preview
                            self.preview = self.scanned.copy()
                            break

                # Convert camera frame for display
                camera_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                
                # Ensure max_width and max_height are valid
                max_width = self.camera_frame.winfo_width()
                max_height = self.camera_frame.winfo_height()
                if max_width > 0 and max_height > 0:
                    camera_frame = self.resize_image_to_fit(camera_frame, max_width, max_height)
                    camera_image = Image.fromarray(camera_frame)
                    camera_photo = ImageTk.PhotoImage(image=camera_image)
                    
                    # Display camera frame in camera label
                    self.camera_label.configure(image=camera_photo)
                    self.camera_label.image = camera_photo

                if self.scanned is not None:
                    # Apply any whitened rectangles
                    if self.whitened_rects:
                        preview_to_show = self.preview.copy()
                        for (start, end) in self.whitened_rects:
                            cv2.rectangle(preview_to_show, start, end, 255, -1)
                    else:
                        preview_to_show = self.preview

                    preview_image = Image.fromarray(preview_to_show)
                    preview_photo = ImageTk.PhotoImage(image=preview_image)
                    self.document_label.configure(image=preview_photo)
                    self.document_label.image = preview_photo

            self.root.after(10, self.update_video)

    def handle_scan(self):
        if self.scanned is not None:
            if not self.scanning:
                self.scanning = True
                self.pdf = FPDF(orientation="P", unit="mm", format="A4")
                self.status_var.set("Starting a new scan session")
                self.speak("Starting a new scan session")
            
            # Apply whitening rectangles before saving
            temp_preview = self.preview.copy()
            if self.whitened_rects:
                for (start, end) in self.whitened_rects:
                    cv2.rectangle(temp_preview, start, end, 255, -1)

            # Save the image with whitening applied
            timestamp = datetime.now().strftime('%H-%M-%S')
            image_path = os.path.join(self.destination_folder, f"temp_{timestamp}.jpg")
            cv2.imwrite(image_path, temp_preview)
        
            # Add to PDF
            self.pdf.add_page()
            self.pdf.image(image_path, x=0, y=0, w=self.A4_width, h=self.A4_height)
            os.remove(image_path)
            
            # Update status
            self.status_var.set(f"Page added to PDF. Total pages: {self.pdf.page_no()}")

    def handle_save(self):
        if self.scanned is not None and self.pdf is not None:
            # Save the PDF
            timestamp = datetime.now().strftime("%m-%d-%y_%H-%M-%S")
            pdf_filename = os.path.join(self.destination_folder, f"scanned_document_{timestamp}.pdf")
            self.pdf.output(pdf_filename, "F")
            
            # Update status and speak
            self.status_var.set("PDF saved successfully")
            self.speak("PDF saved successfully")
            
            # Reset PDF object
            self.pdf = None
            self.scanning = False  # Reset scanning state
        else:
            self.status_var.set("No document scanned yet or no pages added")
            self.speak("No document scanned yet or no pages added")

    def handle_edit(self):
        if self.scanned is not None:
            self.selecting = not self.selecting
            if self.selecting:
                self.status_var.set("Modify mode activated")
                self.speak("Modify mode activated. Select the portion that you want to exclude in scanning.")
                self.whitened_rects = []  # Reset whitened rectangles
                self.preview = self.scanned.copy()
                
                # Bind mouse events for editing
                self.document_label.bind('<Button-1>', self.start_rect)
                self.document_label.bind('<B1-Motion>', self.draw_rect)
                self.document_label.bind('<ButtonRelease-1>', self.end_rect)
            else:
                self.status_var.set("Modify mode deactivated")
                self.speak("Modify mode deactivated.")
                self.whitened_rects = []  # Clear the list of whitened rectangles when exiting erase mode
                self.preview = self.scanned.copy()
                
                # Unbind mouse events
                self.document_label.unbind('<Button-1>')
                self.document_label.unbind('<B1-Motion>')
                self.document_label.unbind('<ButtonRelease-1>')

    def start_rect(self, event):
        if self.selecting:
            self.drawing = True
            self.rect_start = (event.x, event.y)
            self.rect_end = (event.x, event.y)

    def draw_rect(self, event):
        if self.selecting and self.drawing:
            self.rect_end = (event.x, event.y)
            
            # Create a copy of the preview to draw temporary rectangle
            temp_preview = self.preview.copy()
            
            # Convert preview to RGB for drawing
            temp_preview_rgb = cv2.cvtColor(temp_preview, cv2.COLOR_GRAY2RGB)
            
            # Draw the rectangle
            x_min, y_min = min(self.rect_start[0], self.rect_end[0]), min(self.rect_start[1], self.rect_end[1])
            x_max, y_max = max(self.rect_start[0], self.rect_end[0]), max(self.rect_start[1], self.rect_end[1])
            cv2.rectangle(temp_preview_rgb, (x_min, y_min), (x_max, y_max), (0, 0, 255), 2)
            
            # Convert back to PhotoImage
            temp_preview_image = Image.fromarray(temp_preview_rgb)
            temp_preview_photo = ImageTk.PhotoImage(image=temp_preview_image)
            
            # Update label
            self.document_label.configure(image=temp_preview_photo)
            self.document_label.image = temp_preview_photo

    def end_rect(self, event):
        if self.selecting and self.drawing:
            self.drawing = False
            self.rect_end = (event.x, event.y)
            
            # Get rectangle coordinates
            x_min, y_min = min(self.rect_start[0], self.rect_end[0]), min(self.rect_start[1], self.rect_end[1])
            x_max, y_max = max(self.rect_start[0], self.rect_end[0]), max(self.rect_start[1], self.rect_end[1])
            
            # Store rectangle coordinates
            self.whitened_rects.append(((x_min, y_min), (x_max, y_max)))
            
            # White out the selected region
            cv2.rectangle(self.preview, (x_min, y_min), (x_max, y_max), 255, -1)
            
            # Update the preview with the new rectangle
            temp_preview = self.preview.copy()
            temp_preview_image = Image.fromarray(temp_preview)
            temp_preview_photo = ImageTk.PhotoImage(temp_preview_image)
            
            # Update label
            self.document_label.configure(image=temp_preview_photo)
            self.document_label.image = temp_preview_photo

    def handle_open_folder(self):
        os.startfile(self.destination_folder)

    def handle_exit(self):
        self.speech_queue.put(None)  # Signal to exit the speech thread
        self.cap.release()
        self.root.quit()

if __name__ == "__main__":
    root = tk.Tk()
    root.state('zoomed') 
    app = ScannerGUI(root)
    root.mainloop()