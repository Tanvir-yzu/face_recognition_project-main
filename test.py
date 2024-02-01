from sklearn.neighbors import KNeighborsClassifier
import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime
from win32com.client import Dispatch

def speak(str1):
    speak = Dispatch(("SAPI.SpVoice"))
    speak.Speak(str1)

video = cv2.VideoCapture(0)
facedetect = cv2.CascadeClassifier('data/haarcascade_frontalface_default.xml')

# Load face data
try:
    with open('data/faces_data.pkl', 'rb') as f:
        FACES = pickle.load(f)
except Exception as e:
    print("Error loading face data:", e)
    FACES = []

# Load labels
try:
    with open('data/names.pkl', 'rb') as w:
        LABELS = pickle.load(w)
except Exception as e:
    print("Error loading labels:", e)
    LABELS = []

# Check the number of samples in FACES and LABELS
print('Number of samples in FACES:', len(FACES))
print('Number of samples in LABELS:', len(LABELS))

if len(FACES) != len(LABELS):
    raise ValueError("Inconsistent number of samples between FACES and LABELS")

imgBackground = cv2.imread("background.png")

COL_NAMES = ['NAME', 'TIME']

knn = KNeighborsClassifier(n_neighbors=5)

# Check if there is any data to fit the classifier
if len(FACES) > 0 and len(LABELS) > 0:
    knn.fit(FACES, LABELS)

attendance_list = []  # create an empty list to store attendance records

while True:
    ret, frame = video.read()
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = facedetect.detectMultiScale(gray, 1.3, 5)

    for (x, y, w, h) in faces:
        crop_img = frame[y:y+h, x:x+w, :]
        resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)
        output = knn.predict(resized_img)
        ts = time.time()
        date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp = datetime.fromtimestamp(ts).strftime("%H:%M:%S")
        exist = os.path.isfile("Attendance/Attendance_" + date + ".csv")
        cv2.rectangle(frame, (x, y), (x+w, y+h), (0, 0, 255), 1)
        cv2.rectangle(frame, (x, y), (x+w, y+h), (50, 50, 255), 2)
        cv2.rectangle(frame, (x, y-40), (x+w, y), (50, 50, 255), -1)
        cv2.putText(frame, str(output[0]), (x, y-15), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 1)
        cv2.rectangle(frame, (x, y), (x+w, y+h), (50, 50, 255), 1)

        attendance_record = [str(output[0]), str(timestamp)]
        attendance_list.append(attendance_record)  # append each attendance record to the list

    imgBackground[162:162 + 480, 55:55 + 640] = frame
    cv2.imshow("Frame", imgBackground)
    k = cv2.waitKey(1)

    if k == ord('o'):
        speak("Attendance Taken..")
        time.sleep(5)
        if exist:
            with open("Attendance/Attendance_" + date + ".csv", "+a") as csvfile:
                writer = csv.writer(csvfile)
                for record in attendance_list:
                    writer.writerow(record)  # write each attendance record to the CSV file
        else:
            with open("Attendance/Attendance_" + date + ".csv", "+a") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(COL_NAMES)
                for record in attendance_list:
                    writer.writerow(record)  # write each attendance record to the CSV file

        attendance_list = []  # clear the list after writing to the file

    if k == ord('q'):
        break

video.release()
cv2.destroyAllWindows()
