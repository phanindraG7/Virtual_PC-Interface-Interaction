from flask import Flask, render_template, request, redirect, url_for
from spire.presentation.common import *
from spire.presentation import *
from cvzone.HandTrackingModule import HandDetector
import cv2
import os
import numpy as np
# from Actual import VolumeControl,VirtualMouse1

import time
# from NeutralMode import NeutralMode1

app = Flask(__name__)

# Define the upload folder for PPT files
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def ppt_to_img(input_pptx_path, output_directory):
        # Create a Presentation object
        presentation = Presentation()
        # Load a PowerPoint presentation
        presentation.LoadFromFile(input_pptx_path)

        # Create the directory if it doesn't exist
        os.makedirs(output_directory, exist_ok=True)

        # Loop through the slides in the presentation
        for i, slide in enumerate(presentation.Slides):
            # Specify the output file name with the directory path
            file_name = os.path.join(output_directory, str(i) + ".png")
            # Save each slide as a PNG image
            image = slide.SaveAsImage()
            image.Save(file_name)
            image.Dispose()

        # Dispose of the presentation object
        presentation.Dispose()

def gesture_control_presentation(output_directory):
        # Parameters
        width, height = 640, 480
        gesture_threshold = 300

        # Camera Setup
        cap = cv2.VideoCapture(0)
        cap.set(3, width)
        cap.set(4, height)

        # Hand Detector
        detector_hand = HandDetector(detectionCon=0.8, maxHands=1)

        # Variables
        img_list = []
        delay = 40
        button_pressed = False
        counter = 0
        draw_mode = False
        img_number = 0
        delay_counter = 0
        annotations = [[]]
        annotation_number = -1
        annotation_start = False
        hs, ws = int(120 * 1), int(213 * 1)  # width and height of the small image
        fingers = []
        finger2_up_start_time = 0

        # Get list of presentation images
        path_images = sorted(os.listdir(output_directory), key=len)
        print(path_images)

        while True:
            # Get image frame
            success, img = cap.read()
            img = cv2.flip(img, 1)
            path_full_image = os.path.join(output_directory, path_images[img_number])
            img_current = cv2.imread(path_full_image)

            # Find the hand and its landmarks
            hands, img = detector_hand.findHands(img)  # with draw
            # Draw Gesture Threshold line
            cv2.line(img, (0, gesture_threshold), (width, gesture_threshold), (0, 255, 0), 10)

            if hands and not button_pressed:  # If hand is detected
                hand = hands[0]
                cx, cy = hand["center"]
                lm_list = hand["lmList"]  # List of 21 Landmark points
                fingers = detector_hand.fingersUp(hand)  # List of which fingers are up
                # If index finger and pinky finger are up
                if fingers == [0, 1, 0, 0, 1]:
                    # If both fingers are up, start the timer
                    if finger2_up_start_time == 0:
                        finger2_up_start_time = time.time()
                    else:
                        # If the time difference is greater than or equal to 3 seconds, terminate the code
                        if time.time() - finger2_up_start_time >= 3:
                            break
                else:
                    # Reset the start time if fingers are not up
                    finger2_up_start_time = 0

                # Constrain values for easier drawing
                x_val = int(np.interp(lm_list[8][0], [width // 2, width], [0, width]))
                y_val = int(np.interp(lm_list[8][1], [150, height-150], [0, height]))
                index_finger = x_val, y_val

                if cy <= gesture_threshold:  # If hand is at the height of the face
                    if fingers == [1, 0, 0, 0, 0]:
                        print("Left")
                        button_pressed = True
                        if img_number > 0:
                            img_number -= 1
                            annotations = [[]]
                            annotation_number = -1
                            annotation_start = False
                    if fingers == [0, 0, 0, 0, 1]:
                        print("Right")
                        button_pressed = True
                        if img_number < len(path_images) - 1:
                            img_number += 1
                            annotations = [[]]
                            annotation_number = -1
                            annotation_start = False

                if fingers == [0, 1, 1, 0, 0]:
                    cv2.circle(img_current, index_finger, 12, (0, 0, 255), cv2.FILLED)

                if fingers == [0, 1, 0, 0, 0]:
                    if not annotation_start:
                        annotation_start = True
                        annotation_number += 1
                        annotations.append([])
                    print(annotation_number)
                    annotations[annotation_number].append(index_finger)
                    cv2.circle(img_current, index_finger, 12, (0, 0, 255), cv2.FILLED)
                else:
                    annotation_start = False

                if fingers == [0, 1, 1, 1, 0]:
                    if annotations:
                        annotations.pop(-1)
                        annotation_number -= 1
                        button_pressed = True
            else:
                annotation_start = False

            if button_pressed:
                counter += 1
                if counter > delay:
                    counter = 0
                    button_pressed = False

            for i, annotation in enumerate(annotations):
                for j in range(len(annotation)):
                    if j != 0:
                        cv2.line(img_current, annotation[j - 1], annotation[j], (0, 0, 200), 12)

            img_small = cv2.resize(img, (ws, hs))
            h, w, _ = img_current.shape
            img_current[0:hs, w - ws: w] = img_small

            cv2.imshow("Slides", img_current)
            # cv2.imshow("Image", img)

            key = cv2.waitKey(1)
            if key == ord('q'):
                break

