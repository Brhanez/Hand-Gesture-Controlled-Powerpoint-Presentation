# Hand-Gesture-Controlled-Powerpoint-Presentation
Hand Gesture Controlled Powerpoint Presentation

      A Python application that lets you control PowerPoint presentations using hand gestures detected via your webcam. Navigate slides, annotate, and interact with your presentation—all touch-free!

Features

      Hand Gesture Navigation: Move forward/backward through slides using simple hand gestures.
      Real-Time Annotation: Draw directly on slides using your index finger.
      Webcam Overlay: See yourself in a corner of your slides while presenting.
      PowerPoint Integration: Load .pptx files and automatically convert slides to images.
      User-Friendly GUI: Select files, adjust gesture sensitivity, and start presentations easily.

Requirements

      - Windows OS (PowerPoint COM automation required)
      - Python 3.7+
      - Microsoft PowerPoint (installed)
      - Webcam

Python Packages
         - opencv-python
         - cvzone
         - numpy
         - python-pptx
         - comtypes
         - tkinter (usually included with Python)

Install all dependencies with:


        pip install opencv-python cvzone numpy python-pptx comtypes

Usage

        Clone the repository: git clone https://github.com/Brhanez/Hand-Gesture-Controlled-Powerpoint-Presentation.git
        cd hand-gesture-presentation

Run the application:

        python [GUI.py]

In the GUI:

    Click Browse to select a .pptx file.
    Adjust the gesture threshold slider if needed.
    Click Start Presentation.
During Presentation:

    Swipe left (index finger): Previous slide
    Swipe right (pinky): Next slide
    Index finger up: Draw/annotate
    Index & middle fingers up: Pointer/highlight
    Three fingers up: Undo last annotation
    Press q: Quit presentation

How It Works
    
    Uses OpenCV and cvzone to detect hand landmarks and recognize gestures.
    Converts PowerPoint slides to images for fast display and annotation.
    Overlays the webcam feed onto the slides for a modern presentation experience.
    All temporary files are cleaned up after use.


