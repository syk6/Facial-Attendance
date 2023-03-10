
#importing the required libraries
import cv2
import face_recognition
import pandas as pd
from datetime import datetime
import os
import pyautogui
import time
#capture the video from default camera 
webcam_video_stream = cv2.VideoCapture(0)

#load the sample images and get the 128 face embeddings from them
vasanthi_image = face_recognition.load_image_file('images/samples/vasanthi.jpg')
vasanthi_face_encodings = face_recognition.face_encodings(vasanthi_image)[0]

valli_image = face_recognition.load_image_file('images/samples/valli.jpg')
valli_face_encodings = face_recognition.face_encodings(valli_image)[0]

upendra_image = face_recognition.load_image_file('images/samples/Upendra.jpg')
upendra_face_encodings = face_recognition.face_encodings(upendra_image)[0]

yogi_image = face_recognition.load_image_file('images/samples/yogi-min.jpg')
yogi_face_encodings = face_recognition.face_encodings(yogi_image)[0]

#save the encodings and the corresponding labels in seperate arrays in the same order
known_face_encodings = [vasanthi_face_encodings, valli_face_encodings, upendra_face_encodings, yogi_face_encodings]
known_face_names = ["vasanthi", "Sri valli", "Upendra kumar","yogi"]

#initialize the array variable to hold all face locations, encodings and names 
all_face_locations = []
all_face_encodings = []
all_face_names = []
attendance_list= []
#loop through every frame in the video
while True:
    #get the current frame from the video stream as an image
    ret,current_frame = webcam_video_stream.read()
    #resize the current frame to 1/4 size to proces faster
    current_frame_small = cv2.resize(current_frame,(0,0),fx=0.25,fy=0.25)
    #detect all faces in the image
    #arguments are image,no_of_times_to_upsample, model
    all_face_locations = face_recognition.face_locations(current_frame_small,number_of_times_to_upsample=1,model='hog')
    
    #detect face encodings for all the faces detected
    all_face_encodings = face_recognition.face_encodings(current_frame_small,all_face_locations)


    #looping through the face locations and the face embeddings
    for current_face_location,current_face_encoding in zip(all_face_locations,all_face_encodings):
        #splitting the tuple to get the four position values of current face
        top_pos,right_pos,bottom_pos,left_pos = current_face_location
        
        #change the position maginitude to fit the actual size video frame
        top_pos = top_pos*4
        right_pos = right_pos*4
        bottom_pos = bottom_pos*4
        left_pos = left_pos*4
        
        #find all the matches and get the list of matches
        all_matches = face_recognition.compare_faces(known_face_encodings, current_face_encoding)
       
        #string to hold the label
        name_of_person = 'Unknown face'
        
        #check if the all_matches have at least one item
        #if yes, get the index number of face that is located in the first index of all_matches
        #get the name corresponding to the index number and save it in name_of_person
        if True in all_matches:
            first_match_index = all_matches.index(True)
            name_of_person = known_face_names[first_match_index]
            if name_of_person not in attendance_list:
                now = datetime.now()
                date_time = now.strftime("%H:%M:%S")
                attendance_list.append(name_of_person)
                attendance_list.append(date_time)
            
        
        #draw rectangle around the face    
        cv2.rectangle(current_frame,(left_pos,top_pos),(right_pos,bottom_pos),(255,0,0),2)
        
        #to display thr name
        font = cv2.FONT_HERSHEY_DUPLEX
        cv2.putText(current_frame, name_of_person, (left_pos,bottom_pos), font, 1, (34,34,178),1)
    
    #display the video
    cv2.imshow("Webcam Video",current_frame)
    
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

webcam_video_stream.release()
cv2.destroyAllWindows()        

def Attendace(arr):
    
    df = pd.DataFrame()
    # Creating two columns
    df['NAME'] = attendance_list[0::2]
    df['TIME'] = attendance_list[1::2]
    df.to_excel(r'C:\Users\91824\Desktop\result.xlsx', index=False)

Attendace(attendance_list)

students = pd.read_excel(r"C:\Users\91824\Desktop\StudDb.xlsx")
attendees = pd.read_excel(r"C:\Users\91824\Desktop\result.xlsx")

left_join_df = attendees.merge(students, how="left", on="NAME")
left_join_df.to_excel(r"C:\Users\91824\Desktop\MailFile.xlsx")

os.startfile(r"C:\Users\91824\Documents\UiPath\MailTest2\mailAuto.xaml")

time.sleep(300)
pyautogui.press("F6")

    





