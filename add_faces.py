import cv2
import pickle
import numpy as np
import os

if not os.path.exists('data/'):
    os.makedirs('data/')

video = cv2.VideoCapture(0)
face_detect = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')

faces_data = []

i=0
aadhar = input("Enter Aadhar Number : ")

framesTotal = 50
capture_after = 2

while True:
    ret, frame = video.read()
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = face_detect.detectMultiScale(gray, 1.3, 5)
    for (x, y, w, h) in faces:
        crop_img = frame[y:y+h, x:x+w]
        resized_img = cv2.resize(crop_img, (50,50))
        if len(faces_data) <= framesTotal and i%capture_after == 0:
            faces_data.append(resized_img)
        i = i+1
        cv2.putText(frame, str(len(faces_data)),(50,50),cv2.FONT_HERSHEY_COMPLEX, 1, (50,50,255), 0)
        cv2.rectangle(frame, (x,y), (x+w, y+h), (50,50,255), 0)
        

    cv2.imshow('frame', frame)
    k = cv2.waitKey(1)
    if k == ord('q') or len(faces_data) >= framesTotal:
        break
    
video.release()
cv2.destroyAllWindows()



faces_data = np.array(faces_data)
faces_data = faces_data.reshape((framesTotal, -1))
print(faces_data)



if 'names.pkl' not in os.listdir('data/'):
    names = [aadhar]*framesTotal
    with open('data/names.pkl', 'wb') as f:
        pickle.dump(names, f)

else:
    with open('data/names.pkl', 'rb') as f:
        names = pickle.load(f)
    names = names+[aadhar]*framesTotal
    with open('data/names.pkl', 'wb') as f:
        pickle.dump(names, f)
        
if 'faces_data.pkl' not in os.listdir('data/'):
    with open('data/faces_data.pkl', 'wb') as f:
        pickle.dump(faces_data, f)

else:
    with open('data/faces_data.pkl', 'rb') as f:
        faces = pickle.load(f)
    faces = np.append(faces, faces_data, axis=0)
    with open('data/faces_data.pkl', 'wb') as f:
        pickle.dump(faces, f)
