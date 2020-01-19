from imutils.video import VideoStream
from imutils.video import FPS
from pandas import DataFrame
from datetime import date
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import messagebox
import win32com.client as win32
import win32ui
import win32con
import argparse
import imutils
import pickle
import time
import cv2
import os

# global variables
nameList = pd.read_excel(r'name_list.xlsx')
nameList = nameList.set_index('Matric No.')
today = date.today().strftime("%d/%m/%Y")
nameList[today] = 0

outlook = win32.Dispatch('outlook.application')

emailString = ''
newEntry = False
nameEntry = ''
matricEntry = ''
emailEntry = ''
path = ''
picCounter = 0
studentCount = 4

# construct the argument parser and parse the arguments
ap = argparse.ArgumentParser()
ap.add_argument("-d", "--detector", required=True,
	help="path to OpenCV's deep learning face detector")
ap.add_argument("-m", "--embedding-model", required=True,
	help="path to OpenCV's deep learning face embedding model")
ap.add_argument("-r", "--recognizer", required=True,
	help="path to model trained to recognize faces")
ap.add_argument("-l", "--le", required=True,
	help="path to label encoder")
ap.add_argument("-c", "--confidence", type=float, default=0.5,
	help="minimum probability to filter weak detections")
args = vars(ap.parse_args())

# load our serialized face detector from disk
print("[INFO] loading face detector...")
protoPath = os.path.sep.join([args["detector"], "deploy.prototxt"])
modelPath = os.path.sep.join([args["detector"],
	"res10_300x300_ssd_iter_140000.caffemodel"])
detector = cv2.dnn.readNetFromCaffe(protoPath, modelPath)

# load our serialized face embedding model from disk
print("[INFO] loading face recognizer...")
embedder = cv2.dnn.readNetFromTorch(args["embedding_model"])

# load the actual face recognition model along with the label encoder
recognizer = pickle.loads(open(args["recognizer"], "rb").read())
le = pickle.loads(open(args["le"], "rb").read())

# initialize the video stream, then allow the camera sensor to warm up
print("[INFO] starting video stream...")
vs = VideoStream(src=0, resolution=(1920, 1080)).start()
time.sleep(2.0)

# start the FPS throughput estimator
fps = FPS().start()

# loop over frames from the video file stream
while True:
	# grab the frame from the threaded video stream
	frame = vs.read()

	# resize the frame to have a width of 600 pixels (while
	# maintaining the aspect ratio), and then grab the image
	# dimensions
	frame = imutils.resize(frame, width=1080)
	(h, w) = frame.shape[:2]

	if newEntry and picCounter > 4:
		nameList.loc[matricEntry] = 0
		nameList.at[matricEntry,'Name'] = nameEntry
		nameList.at[matricEntry,'Email'] = emailEntry
		newEntry = False

	# construct a blob from the image
	imageBlob = cv2.dnn.blobFromImage(
		cv2.resize(frame, (300, 300)), 1.0, (300, 300),
		(104.0, 177.0, 123.0), swapRB=False, crop=False)

	# apply OpenCV's deep learning-based face detector to localize
	# faces in the input image
	detector.setInput(imageBlob)
	detections = detector.forward()

	# loop over the detections
	for i in range(0, detections.shape[2]):
		# extract the confidence (i.e., probability) associated with
		# the prediction
		confidence = detections[0, 0, i, 2]

		# filter out weak detections
		if confidence > args["confidence"]:
			# compute the (x, y)-coordinates of the bounding box for
			# the face
			box = detections[0, 0, i, 3:7] * np.array([w, h, w, h])
			(startX, startY, endX, endY) = box.astype("int")

			# extract the face ROI
			face = frame[startY:endY, startX:endX]
			(fH, fW) = face.shape[:2]

			# ensure the face width and height are sufficiently large
			if fW < 20 or fH < 20:
				continue

			# construct a blob for the face ROI, then pass the blob
			# through our face embedding model to obtain the 128-d
			# quantification of the face
			faceBlob = cv2.dnn.blobFromImage(face, 1.0 / 255,
				(96, 96), (0, 0, 0), swapRB=True, crop=False)
			embedder.setInput(faceBlob)
			vec = embedder.forward()

			# perform classification to recognize the face
			preds = recognizer.predict_proba(vec)[0]
			j = np.argmax(preds)
			proba = preds[j]
			name = le.classes_[j]
			studentName = nameList.loc[name]['Name']

			# draw the bounding box of the face along with the
			# associated probability
			text = "{}: {:.2f}%".format(studentName, proba * 100)
			y = startY - 10 if startY - 10 > 10 else startY + 10

			cv2.rectangle(frame, (startX, startY), (endX, endY), (100, 100, 100), 2)
			cv2.putText(frame, text, (startX, y),
				cv2.FONT_HERSHEY_SIMPLEX, 0.60, (100, 100, 100), 2)
			if proba >= 0.7:
				cv2.rectangle(frame, (startX, startY), (endX, endY), (158, 98, 226), 2)
				cv2.putText(frame, text, (startX, y), cv2.FONT_HERSHEY_SIMPLEX, 
					0.60, (158, 98, 226), 2)
				if nameList.loc[name][today] == 0:
					nameList.at[name, today] = 1
					emailString = emailString + nameList.loc[name]['Email'] + '; '


	# update the FPS counter
	fps.update()

	# show the output frame
	cv2.imshow("iSeeU", frame)
	key = cv2.waitKey(1) & 0xFF

	# if the `q` key was pressed, break from the loop
	if key == ord("q"):
		break

	# if the `a` key was pressed, a picture will be captured
	if key == ord("a"):
		cv2.imwrite(os.path.join(path, '0000' + str(picCounter) + '.png'), frame)
		
		if picCounter >= 4:
			print("Congratulations!\nYou are done!")
		else:
			print("Take pictures!\nYou have " + str(4 - picCounter) + " pictures to go!")
		picCounter = picCounter + 1

	# if the `s` key was pressed, add a new student into the database
	if key == ord("s"):
		def save_entry_fields():
			global nameEntry
			global matricEntry
			global emailEntry
			global newEntry
			global picCounter
			global path

			matricEntry = e1.get()
			nameEntry = e2.get()
			emailEntry = e3.get()
			picCounter = 0
			newEntry = True
			
			messagebox.showinfo(title="Take pictures!", message="Click 'A' to take picture!")
			#path = 'C:/Users/syiny/Desktop/live-face-recognition/dataset/' + matricEntry
			path = 'C:/Users/ACER/Hackathon/live-face-recognition/dataset/' + matricEntry
			os.mkdir(path)

			master.destroy()

		master = tk.Tk()
		tk.Label(master, 
		         text="Matric No.", font=20).grid(row=0)
		tk.Label(master, 
		         text="Name", font=20).grid(row=1)
		tk.Label(master, 
		         text="Email", font=20).grid(row=2)

		e1 = tk.Entry(master)
		e2 = tk.Entry(master)
		e3 = tk.Entry(master)

		e1.grid(row=0, column=1)
		e2.grid(row=1, column=1)
		e3.grid(row=2, column=1)


		tk.Button(master, 
		          text='Cancel', 
		          command=master.destroy).grid(row=4, 
		                                    column=0, 
		                                    sticky=tk.W, 
		                                    pady=4)
		tk.Button(master, 
		          text='Enter', command=save_entry_fields).grid(row=4, 
		                                                       column=1, 
		                                                       sticky=tk.W, 
		                                                       pady=4)

		tk.mainloop()

writer = pd.ExcelWriter('name_list.xlsx', engine='xlsxwriter')
nameList.to_excel(writer, sheet_name='Attendance')
writer.save()

# sends an email to those who have attended
if emailString != '':
	mail = outlook.CreateItem(0)
	mail.To = emailString
	mail.Subject = 'Confirmation of Attendance'
	mail.Body = 'Your presence is greatly appreciated and has been noted on ' + today + '! :)'
	mail.Send()

# stop the timer and display FPS information
fps.stop()
print("[INFO] elasped time: {:.2f}".format(fps.elapsed()))
print("[INFO] approx. FPS: {:.2f}".format(fps.fps()))

# do a bit of cleanup
cv2.destroyAllWindows()
vs.stop()
