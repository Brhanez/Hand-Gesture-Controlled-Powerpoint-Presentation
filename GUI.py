import cv2
import os
import numpy as np
from cvzone.HandTrackingModule import HandDetector
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pptx import Presentation
import comtypes.client
import uuid
import shutil

# Parameters
width, height = 1280, 720
gestureThreshold = 300
outputFolder = "SlideImages"
hs, ws = int(120 * 1), int(213 * 1)  # width and height of small image

class PresentationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PowerPoint Presentation Viewer")
        self.root.geometry("400x300")

        # GUI Elements
        self.label = tk.Label(root, text="Select a PowerPoint presentation file")
        self.label.pack(pady=10)

        # Gesture threshold slider
        self.threshold_label = tk.Label(root, text=f"Gesture Threshold: {gestureThreshold}")
        self.threshold_label.pack(pady=5)
        self.threshold_slider = ttk.Scale(root, from_=100, to_=500, orient=tk.HORIZONTAL, command=self.update_threshold)
        self.threshold_slider.set(gestureThreshold)
        self.threshold_slider.pack(pady=5)

        self.browse_button = tk.Button(root, text="Browse", command=self.browse_file)
        self.browse_button.pack(pady=10)

        self.start_button = tk.Button(root, text="Start Presentation", command=self.start_presentation, state=tk.DISABLED)
        self.start_button.pack(pady=10)

        self.quit_button = tk.Button(root, text="Quit", command=self.quit_app)
        self.quit_button.pack(pady=10)

        self.pptx_file = None
        self.image_folder = None
        self.gesture_threshold = gestureThreshold

    def update_threshold(self, value):
        self.gesture_threshold = int(float(value))
        self.threshold_label.config(text=f"Gesture Threshold: {self.gesture_threshold}")

    def browse_file(self):
        self.pptx_file = filedialog.askopenfilename(
            title="Select PowerPoint File",
            filetypes=[("PowerPoint files", "*.pptx *.ppt")]
        )
        if self.pptx_file:
            # Verify file is a valid PowerPoint file
            try:
                prs = Presentation(self.pptx_file)
                slide_count = len(prs.slides)
                if slide_count > 0:
                    self.label.config(text=f"Selected: {os.path.basename(self.pptx_file)} ({slide_count} slides)")
                    self.start_button.config(state=tk.NORMAL)
                    self.extract_slides_to_png()
                else:
                    messagebox.showerror("Error", "No slides found in the selected presentation")
                    self.start_button.config(state=tk.DISABLED)
            except Exception as e:
                messagebox.showerror("Error", f"Invalid PowerPoint file: {str(e)}")
                self.start_button.config(state=tk.DISABLED)
        else:
            self.start_button.config(state=tk.DISABLED)

    def extract_slides_to_png(self):
        try:
            # Ensure the PowerPoint file path is absolute and properly formatted
            pptx_file_abs = os.path.abspath(self.pptx_file)
            if not os.path.exists(pptx_file_abs):
                raise FileNotFoundError(f"PowerPoint file not found: {pptx_file_abs}")

            # Create a unique temporary folder for slide images
            unique_id = str(uuid.uuid4())
            self.image_folder = os.path.join(os.getcwd(), outputFolder, unique_id)
            os.makedirs(self.image_folder, exist_ok=True)

            # Initialize PowerPoint COM object with better error handling
            try:
                comtypes.CoInitialize()  # Initialize COM library
                powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
            except Exception as e:
                raise RuntimeError(f"Failed to initialize PowerPoint COM object. Ensure PowerPoint is installed: {str(e)}")

            try:
                # Ensure PowerPoint is visible and try opening the presentation
                powerpoint.Visible = 1
                presentation = powerpoint.Presentations.Open(pptx_file_abs)
            except Exception as e:
                powerpoint.Quit()
                raise RuntimeError(f"Failed to open PowerPoint presentation: {str(e)}")

            # Export slides as images in order
            try:
                slide_count = len(presentation.Slides)
                for i in range(1, slide_count + 1):  # Explicitly iterate from 1 to slide count
                    slide_path = os.path.join(self.image_folder, f"slide{i}.png")
                    presentation.Slides(i).Export(slide_path, "PNG", 1280, 720)
            except Exception as e:
                raise RuntimeError(f"Failed to export slides: {str(e)}")
            finally:
                # Always close the presentation and PowerPoint
                presentation.Close()
                powerpoint.Quit()
                comtypes.CoUninitialize()  # Uninitialize COM library

            # Verify and rename files to ensure order
            image_files = [f for f in os.listdir(self.image_folder) if f.lower().endswith('.png')]
            if len(image_files) != slide_count:
                raise RuntimeError(f"Expected {slide_count} slides, but found {len(image_files)} images.")

            for idx, old_name in enumerate(sorted(image_files, key=lambda x: int(x.split('slide')[1].split('.png')[0])), 1):
                new_name = f"{idx}.png"
                os.rename(
                    os.path.join(self.image_folder, old_name),
                    os.path.join(self.image_folder, new_name)
                )

            global folderPath
            folderPath = self.image_folder  # Update global folderPath to extracted images

        except Exception as e:
            messagebox.showerror("Error", f"Failed to extract slides: {str(e)}")
            self.image_folder = None
            self.start_button.config(state=tk.DISABLED)
            if os.path.exists(self.image_folder):
                shutil.rmtree(self.image_folder)

    def quit_app(self):
        if self.image_folder and os.path.exists(self.image_folder):
            shutil.rmtree(self.image_folder)  # Clean up temporary folder
        self.root.destroy()

    def start_presentation(self):
        if not self.image_folder or not os.path.exists(self.image_folder):
            messagebox.showerror("Error", "No slides extracted. Please select a valid PowerPoint file.")
            return
        self.root.destroy()  # Close GUI and start presentation
        self.run_presentation()

    def run_presentation(self):
        # Camera Setup
        cap = cv2.VideoCapture(0)
        cap.set(3, width)
        cap.set(4, height)

        # Hand Detector
        detectorHand = HandDetector(detectionCon=0.8, maxHands=1)

        # Variables
        imgList = []
        delay = 30
        buttonPressed = False
        counter = 0
        drawMode = False
        imgNumber = 0
        delayCounter = 0
        annotations = [[]]
        annotationNumber = -1
        annotationStart = False

        # Get list of presentation images
        pathImages = sorted([f for f in os.listdir(folderPath) if f.lower().endswith('.png')], key=lambda x: int(os.path.splitext(x)[0]))
        print(pathImages)

        while True:
            # Get image frame
            success, img = cap.read()
            if not success:
                print("Failed to capture image from webcam")
                break
            img = cv2.flip(img, 1)
            pathFullImage = os.path.join(folderPath, pathImages[imgNumber])
            imgCurrent = cv2.imread(pathFullImage)
            if imgCurrent is None:
                print(f"Failed to load image: {pathFullImage}")
                break

            # Create an overlay for the semi-transparent line
            overlay = img.copy()
            cv2.line(overlay, (0, self.gesture_threshold), (width, self.gesture_threshold), (0, 255, 0), 10)
            alpha = 0.3  # Transparency factor
            img = cv2.addWeighted(overlay, alpha, img, 1 - alpha, 0)

            # Find the hand and its landmarks
            hands, img = detectorHand.findHands(img)  # with draw

            if hands and buttonPressed is False:  # If hand is detected
                hand = hands[0]
                cx, cy = hand["center"]
                lmList = hand["lmList"]  # List of 21 Landmark points
                fingers = detectorHand.fingersUp(hand)  # List of which fingers are up

                # Constrain values for easier drawing
                xVal = int(np.interp(lmList[8][0], [width // 2, width], [0, width]))
                yVal = int(np.interp(lmList[8][1], [150, height-150], [0, height]))
                indexFinger = xVal, yVal

                if cy <= self.gesture_threshold:  # If hand is at or above the threshold
                    if fingers == [1, 0, 0, 0, 0]:
                        print("Left")
                        buttonPressed = True
                        if imgNumber > 0:
                            imgNumber -= 1
                            annotations = [[]]
                            annotationNumber = -1
                            annotationStart = False
                    if fingers == [0, 0, 0, 0, 1]:
                        print("Right")
                        buttonPressed = True
                        if imgNumber < len(pathImages) - 1:
                            imgNumber += 1
                            annotations = [[]]
                            annotationNumber = -1
                            annotationStart = False

                if fingers == [0, 1, 1, 0, 0]:
                    cv2.circle(imgCurrent, indexFinger, 12, (0, 0, 255), cv2.FILLED)

                if fingers == [0, 1, 0, 0, 0]:
                    if annotationStart is False:
                        annotationStart = True
                        annotationNumber += 1
                        annotations.append([])
                    print(annotationNumber)
                    annotations[annotationNumber].append(indexFinger)
                    cv2.circle(imgCurrent, indexFinger, 12, (0, 0, 255), cv2.FILLED)

                else:
                    annotationStart = False

                if fingers == [0, 1, 1, 1, 0]:
                    if annotations:
                        annotations.pop(-1)
                        annotationNumber -= 1
                        buttonPressed = True

            else:
                annotationStart = False

            if buttonPressed:
                counter += 1
                if counter > delay:
                    counter = 0
                    buttonPressed = False

            # Draw annotations
            for i, annotation in enumerate(annotations):
                for j in range(len(annotation)):
                    if j != 0:
                        cv2.line(imgCurrent, annotation[j - 1], annotation[j], (0, 0, 200), 12)

            # Display slide number and total slides
            slide_text = f"Slide {imgNumber + 1}/{len(pathImages)}"
            cv2.putText(imgCurrent, slide_text, (10, 50), cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 255, 255), 2, cv2.LINE_AA)

            # Overlay webcam feed
            imgSmall = cv2.resize(img, (ws, hs))
            h, w, _ = imgCurrent.shape
            imgCurrent[0:hs, w - ws: w] = imgSmall

            cv2.imshow("Slides", imgCurrent)
            cv2.imshow("Image", img)

            key = cv2.waitKey(1)
            if key == ord('q'):
                cap.release()
                cv2.destroyAllWindows()
                break

        # Clean up
        cap.release()
        cv2.destroyAllWindows()
        if self.image_folder and os.path.exists(self.image_folder):
            shutil.rmtree(self.image_folder)  # Remove temporary slide images

if __name__ == "__main__":
    root = tk.Tk()
    app = PresentationApp(root)
    root.mainloop()