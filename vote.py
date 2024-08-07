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
    speak=Dispatch(("SAPI.SpVoice"))
    speak.Speak(str1)
    
video = cv2.VideoCapture(0)
face_detect = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')

if not os.path.exists('data/'):
    os.makedirs('data/') 

with open('data/names.pkl', 'rb') as f:
    LABELS = pickle.load(f)
    
with open('data/faces_data.pkl', 'rb') as f:
    FACES = pickle.load(f)
    
knn = KNeighborsClassifier(n_neighbors = 5)

knn.fit(FACES, LABELS)

imgBackground = cv2.imread("background.png")

COL_NAMES = ['AADHAR', 'VOTE', 'DATE', 'TIME']

while True:
    ret, frame = video.read()
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = face_detect.detectMultiScale(gray, 1.3, 5)
    for (x, y, w, h) in faces:
        crop_img = frame[y:y+h, x:x+w]
        resized_img = cv2.resize(crop_img, (50,50)).flatten().reshape(1,-1)
        output = knn.predict(resized_img)
        ts = time.time()
        date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp = datetime.fromtimestamp(ts).strftime("%H:%M-%S")
        exist = os.path.isfile("Votes" + ".csv")
        cv2.rectangle(frame, (x,y), (x+w, y+h), (0,0,255), 1)
        cv2.rectangle(frame, (x,y), (x+w, y+h), (50,50,255), 2)
        cv2.rectangle(frame, (x,y-40), (x+w, y), (50,50,255), -1)
        cv2.putText(frame, str(output[0]), (x,y-15), cv2.FONT_HERSHEY_COMPLEX, 1, (255,255,255), 1)
        cv2.rectangle(frame, (x,y), (x+w, y+h), (50,50,255), 1)
        attendance = [output[0], timestamp]
    
    imgBackground[370:370 + 480, 255:255 + 640] = frame
        
        
    cv2.imshow('frame',imgBackground)
    k = cv2.waitKey(1)
    
    def check_exist(value):
        try:
            with open("Votes.csv", "r") as csvfile:
                reader = csv.reader(csvfile)
                for row in reader:
                    if row and row[0] == value:
                        return True
        except FileNotFoundError:
            print("File Not Found or Unable to open")
        
        return False
    voter_exist = check_exist(output[0])
    if voter_exist:
        speak("You have already voted")
        break
    if k==ord('1'):
        speak("Your vote has been recorded")
        time.sleep(5)
        if exist:
            with open("Votes" + ".csv", "+a") as csvfile:
                writer = csv.writer(csvfile)
                attendance = [output[0],"BJP", date, timestamp]
                writer.writerow(attendance)
            csvfile.close()
        else:
            with open("Votes" + ".csv", "+a") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(COL_NAMES)
                attendance = [output[0],"BJP", date, timestamp]
                writer.writerow(attendance)
            csvfile.close()
        speak("THANK YOU FOR PARTICIPATION IN ELECTION")


    if k==ord('2'):
        speak("Your vote has been recorded")
        time.sleep(5)
        if exist:
            with open("Votes" + ".csv", "+a") as csvfile:
                writer = csv.writer(csvfile)
                attendance = [output[0],"CONGRESS", date, timestamp]
                writer.writerow(attendance)
            csvfile.close()
        else:
            with open("Votes" + ".csv", "+a") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(COL_NAMES)
                attendance = [output[0], "CONGRESS", date, timestamp]
                writer.writerow(attendance)
            csvfile.close()
        speak("THANK YOU FOR PARTICIPATION IN ELECTION")
        
        
    if k==ord('3'):
        speak("Your vote has been recorded")
        time.sleep(5)
        if exist:
            with open("Votes" + ".csv", "+a") as csvfile:
                writer = csv.writer(csvfile)
                attendance = [output[0], "AAP", date, timestamp]
                writer.writerow(attendance)
            csvfile.close()
        else:
            with open("Votes" + ".csv", "+a") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(COL_NAMES)
                attendance = [output[0],"AAP", date, timestamp]
                writer.writerow(attendance)
            csvfile.close()
        speak("THANK YOU FOR PARTICIPATION IN ELECTION")
        
        

    if k==ord('4'):
        speak("Your vote has been recorded")
        time.sleep(5)
        if exist:
            with open("Votes" + ".csv", "+a") as csvfile:
                writer = csv.writer(csvfile)
                attendance = [output[0], "NOTA", date, timestamp]
                writer.writerow(attendance)
            csvfile.close()
        else:
            with open("Votes" + ".csv", "+a") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(COL_NAMES)
                attendance = [output[0], "NOTA", date, timestamp]
                writer.writerow(attendance)
            csvfile.close()
        speak("THANK YOU FOR PARTICIPATION IN ELECTION")


video.release()
cv2.destroyAllWindows()    