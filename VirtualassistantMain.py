from PyQt5.QtWidgets import QWidget
from PyQt5.QtWidgets import QApplication
from PyQt5 import QtGui
from PyQt5.QtCore import QThread
from VirtualassistantUi import Ui_Form
from PyQt5.QtCore import QTimer,QTime,QDate
import sys
import pyttsx3
import speech_recognition as sr
import datetime
from time import sleep
import os
from requests import get
import wikipedia
import webbrowser
import pyjokes
import pywhatkit
import pyautogui
import wolframalpha
import requests
import json
import datetime
from googletrans import Translator
from gtts import gTTS
from playsound import playsound
from pywikihow import search_wikihow
import openai
import PyPDF2
import tkfilebrowser as fb
import cv2
from win32api import GetSystemMetrics
import numpy as np
import time



engine=pyttsx3.init("sapi5")
voices=engine.getProperty("voices")
#print(voice[1].id)
engine.setProperty("voice",voices[0].id)

#Text to speech
def speak(audio):
    virtualassistant_gui.updatemoviesdynamically("Speaking")
    engine.say(audio)
    virtualassistant_gui.terminalprint(audio)
    engine.runAndWait()

#To convert voice into text
def takecommand():

    r=sr.Recognizer()
    with sr.Microphone() as source:
        virtualassistant_gui.updatemoviesdynamically("Listening")
        virtualassistant_gui.terminalprint("Listening...")
        r.pause_threshold=1
        audio=r.listen(source)

    try:
        virtualassistant_gui.updatemoviesdynamically("Loading")
        virtualassistant_gui.terminalprint("Recognizing...")
        query=r.recognize_google(audio, language="en-in")
        virtualassistant_gui.terminalprint(f"User said:{query}\n")
    
    except Exception as e:
        virtualassistant_gui.terminalprint("Say that again please...")
        return "none"
    return query

#To wish
def wish():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!")

    elif hour>=12 and hour<18:
        speak("Good Afternoon!")   

    else:
        speak("Good Evening!")  

    speak("I am Mark Sir. Please tell me how may I help you")

def Wolframe(query):
    api_key="QV7KUW-Y7VU6YH239"
    requester=wolframalpha.Client(api_key)
    requested=requester.query(query)

    try:
        Answer=next(requested.results).text
        return Answer
    
    except:
        speak("The value is not answerable")

def Calculator(query):
    Term=str(query)
    Term=Term.replace("mark","")
    Term=Term.replace("plus","+")
    Term=Term.replace("minus","-")
    Term=Term.replace("multiply","*")
    Term=Term.replace("into","*")
    Term=Term.replace("divided by","/")

    Final=str(Term)

    try:
        result=Wolframe(Final)
        speak(f"{result}")

    except:
        speak("The value is not answerable")

def latestnews():
    api_dict={"business" : "https://newsapi.org/v2/top-headlines?country=in&category=business&apiKey=c5e1cb68de73429fa36213cc5247533b",
         "entertainment" : "https://newsapi.org/v2/top-headlines?country=in&category=entertainment&apiKey=c5e1cb68de73429fa36213cc5247533b",
                "health" : "https://newsapi.org/v2/top-headlines?country=in&category=health&apiKey=c5e1cb68de73429fa36213cc5247533b",
               "science" : "https://newsapi.org/v2/top-headlines?country=in&category=science&apiKey=c5e1cb68de73429fa36213cc5247533b",
            "technology" : "https://newsapi.org/v2/top-headlines?country=in&category=technology&apiKey=c5e1cb68de73429fa36213cc5247533b",
                "sports" : "https://newsapi.org/v2/top-headlines?country=in&category=sports&apiKey=c5e1cb68de73429fa36213cc5247533b"}

    content= None
    url= None
    speak("Which field news do you want, [business], [entertainment] , [health] , [science] , [technology] , [sports]")
    field = takecommand().lower()
    try:
        for key, value in api_dict.items():
            if key.lower() in field.lower():
                url=value
            

        news=requests.get(url).text
        news=json.loads(news)
        speak("Here is the first news")

        arts=news["articles"]
        for articles in arts:
            article=articles["title"]
            speak(article)
            
            speak("Say yes to continue listening more news and no to stop")
            a=takecommand().lower()
            if str(a)=="yes":
                pass
            elif str(a)=="no":
                speak("Thats all")
                break
    except:
        speak("Please choose the correct field")

def translategl(query):
    speak("Sure sir")
    translator=Translator()
    speak("Choose the language in which you want to translate, [Hindi], [Marathi], [Bengali], [Spanish], [French], [Arabic], [Russian], [Portuguese], [Japanese], [Korean]")
    
    b=takecommand().lower()
    b=b.replace("hindi","hi")
    b=b.replace("marathi","mr")
    b=b.replace("bengali","bn")
    b=b.replace("spanish","es")
    b=b.replace("french","fr")
    b=b.replace("arabic","ar")
    b=b.replace("russian","ru")
    b=b.replace("portuguese","pt")
    b=b.replace("japanese","ja")
    b=b.replace("korean","ko")
    b1=str(b)

    text_to_translate=translator.translate(query,src="auto",dest=b1)
    text=text_to_translate.text
    try:
        speakg1=gTTS(text=text,lang=b,slow=False)
        speakg1.save("voice.mp3")
        playsound("voice.mp3")
        sleep(5)
        os.remove("voice.mp3")
    except:
        speak("Unable to translate")

def Query(query):
    term=str(query)
    term=term.replace("mark","")
    final=str(term)

    try:
        result=Wolframe(final)
        speak(f"{result}")

    except:
        speak("Try again later")

def writenote():
    speak("What should I write, sir")
    note=takecommand()
    file = open('Notes.txt', 'a')
    time=datetime.datetime.now().strftime("%d/%B/%y")
    file.write(time)
    file.write(" --> ")
    file.write(note)
    file.write("\n")
    speak("Point noted successfully.")

dictapp = {"command prompt":"cmd","paint":"mspaint","word":"winword","excel":"excel","chrome":"chrome","vscode":"code","powerpoint":"powerpnt"}
def openapp(query):
    query=query.replace("open","")
    query=query.replace("mark","")
    keys=list(dictapp.keys())
    for app in keys:
        if app in query:
            os.system(f"start {dictapp[app]}")

def closeapp(query):
    query=query.replace("close","")
    query=query.replace("mark","")
    keys=list(dictapp.keys())
    for app in keys:
        if app in query:
            speak("Ok sir, closing"+query)
            os.system(f"taskkill /f /im {dictapp[app]}.exe") 

def Speedtest():
    import speedtest
    speak("Checking speed")
    speed=speedtest.Speedtest()
    downloading=speed.download()
    correctdownload=int(downloading/1048576)
    uploading=speed.upload()
    correctupload=int(uploading/1048576)
    speak(f"The downloading speed is {correctdownload} mbp s and the uploading speed is {correctupload} mbp s")


def Message():
    strhour=int(datetime.datetime.now().strftime("%H"))
    strtime=int(datetime.datetime.now().strftime("%M"))
    strtime=strtime+2

    speak("Speak the number to whom do you want to send the message")
    recipent=takecommand()
    recipent=recipent.replace(" ","")
    recipent="+91"+recipent

    speak("Whats the message")
    message=takecommand()
    pywhatkit.sendwhatmsg(recipent,message,strhour,strtime)

def search(query):
    max_results=1
    command=search_wikihow(query,max_results)
    assert len(command)==1
    speak(command[0].summary)

def Email():
    speak("Speak the email address of the receiver")
    receiver=takecommand().lower()
    receiver=receiver.lower()
    receiver=receiver.replace("at the rate","@")
    receiver=receiver.replace(" ","")
    virtualassistant_gui.terminalprint(receiver)

    speak("What is the subject of the email")
    subject=takecommand()
    virtualassistant_gui.terminalprint(subject)
    
    speak("Speak the message that is to be mailed")
    message=takecommand()
    pywhatkit.send_mail("atishp590@gmail.com","vzqkwzwlkisyadhd",subject,message,receiver)
    speak("Email sent successfully")

def Reply(question,chat_log=None):
    openai.api_key="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"
    completion=openai.Completion()

    chat_log_template='''You : Hello, who are you?
    Mark : I am doing great. How can I help you?
    You : Who have developed you?
    Mark : I am developed by atish '''
    
    if chat_log is None:
        chat_log=chat_log_template
    prompt=f"{chat_log}You : {question}\n Mark :"
    response=completion.create(
        prompt=prompt, engine="text-davinci-003", stop=["\nYou"], temperature=0.9,
        top_p=1, frequency_penalty=0, presence_penalty=0.6, best_of=1,
        max_tokens=150)
    answer=response.choices[0].text.strip()
    speak(answer)

def ai(prompt):
    openai.api_key = "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"
    text = f"OpenAI response for Prompt: {prompt} \n *************************\n\n"

    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=prompt,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    # print(response["choices"][0]["text"])
    try:
        text += response["choices"][0]["text"]
        if not os.path.exists("Openai"):
            os.mkdir("Openai")

        # with open(f"Openai/prompt- {random.randint(1, 2343434356)}", "w") as f:
        with open(f"Openai/{''.join(prompt.split('intelligence')[1:]).strip() }.txt", "w") as f:
            f.write(text)
        speak("Task completed")
    except:
        speak("Try again later")

def Pdfread():
    speak("Please select the pdf file")
    
    pdf=fb.askopenfilename()
    
    pdfreader=PyPDF2.PdfReader(pdf)
    pages=len(pdfreader.pages)
    speak(f"Number of pages in this pdf are {pages}")
    speak("Which page number I should read")
    
    pagenumber=takecommand().lower()
    pagenumber=pagenumber.replace("page number","")
    numpage=int(pagenumber)
    numpage=numpage-1
    page=pdfreader.pages[numpage]
    
    text=page.extract_text()
    speak("Choose the language in which I have to read [Hindi], [English]")
    language=takecommand().lower()

    if "hindi" in language:
        transl=Translator()
        txtHindi=transl.translate(text,"hi")
        translatedtxt=txtHindi.text
        speech=gTTS(text=translatedtxt)
        
        try:
            speech.save("pdf.mp3")
            virtualassistant_gui.updatemoviesdynamically("Speaking")
            playsound("pdf.mp3")
            os.remove("pdf.mp3")
        except:
            speak("Try again later")
    else:
        speak(text)

def Screenrec():
    width=GetSystemMetrics(0)
    height=GetSystemMetrics(1)

    dimension=(width,height)

    f=cv2.VideoWriter_fourcc("m","p","4","v")

    file=r"C:\Users\Aniket Patil\Videos\Captures\video.mp4"

    output=cv2.VideoWriter(file,f,30.0,dimension)

    speak("For how many minutes the screen should be recorded")
    minutes=takecommand()
    minutes=minutes.replace("minutes","")
    minutes=minutes.replace("minute","")
    minutes=int(minutes)
    minutes=minutes*60

    starttime=time.time()
    duration=minutes+60
    endtime=starttime+duration
    while True:
        image=pyautogui.screenshot()
        frame1=np.array(image)
        frame=cv2.cvtColor(frame1,cv2.COLOR_BGR2RGB)
        output.write(frame)
        ctime=time.time()
        virtualassistant_gui.terminalprint(ctime)
        if ctime > endtime:
            break
    output.release()
    speak("Screen recording has completed")

class VirtualassistantMain(QThread):
    def __init__(self):
        super(VirtualassistantMain,self).__init__()

    def run(self):
        self.taskexecution()

    def taskexecution(self):
        while True:
            command=takecommand()
            if "hello mark" in command:
                wish()
                while True:
                    query=takecommand().lower()

                    #Logic for executing task based on query

                    if "ip address" in query:
                        ip=get("https://api.ipify.org").text
                        speak(f"Your ip address is {ip}")   

                    elif "wikipedia" in query:
                        speak("Searching wikipedia...")
                        query=query.replace("Wikipedia","")
                        results=wikipedia.summary(query,sentences=2) 
                        speak("According to wikipedia")
                        speak(results)    

                    elif "open google" in query:
                        speak("Sir, what should I search on google")
                        cm=takecommand().lower()
                        webbrowser.open(f"{cm}")
                    
                    #Commands to open any website
                    elif "visit" in query:
                        name=query.replace("visit","")
                        NameA=str(name)

                        if "youtube" in NameA:
                            webbrowser.open("https://www.youtube.com/")

                        elif "instagram" in NameA:
                            webbrowser.open("https://www.instagram.com/")

                        else:
                            string=("https://www."+NameA+".com")
                            string2=string.replace(" ","")
                            webbrowser.open(string2)

                    #To find jokes
                    elif "joke" in query:
                        joke=pyjokes.get_joke("en",category="neutral")
                        speak(joke)

                    elif "where is" in query:
                        query = query.replace("where is", "")
                        location = query
                        speak("User asked to Locate")
                        speak(location)
                        webbrowser.open("https://www.google.nl / maps / place/" + location)

                    #Commands to shutdown,restart and sleep the system
                    elif "shutdown the system" in query:
                        os.system("shutdown /s /t S")

                    elif "restart" in query:
                        os.system("shutdown /r /t S")

                    elif "sleep" in query:
                        os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")

                    #Commands to control volume of system
                    elif "volume up" in query:
                        pyautogui.press("volumeup")

                    elif "volume down" in query:
                        pyautogui.press("volumedown")

                    elif "volume mute" in query:
                        pyautogui.press("volume.mute")

                    #Command to play song
                    elif "song" in query:
                        speak("Sir, which song do you want to play")
                        term=takecommand()
                        result="https://www.youtube.com/results?search_query=" + term
                        webbrowser.open(result)
                        pywhatkit.playonyt(term)
                        speak("Playing song"+term)    
            
                    #Command to take screenshot
                    elif "screenshot" in query:
                        speak("Sir, please tell me the name for this screenshot file")
                        name=takecommand().lower()
                        speak("Sir, please hold the screen for few seconds I am taking a screenshot")
                        sleep(3)
                        img=pyautogui.screenshot()
                        img.save(f"{name}.png")
                        speak("I am done Sir, the screenshot is saved in our main folder")   
                    
                    #Command to perform calculation
                    elif "calculate" in query:
                        query=query.replace("calculate","")
                        query=query.replace("mark","")
                        Calculator(query)
                    
                    #command to get news
                    elif "news" in query:
                        latestnews()
                    
                    #command to translate
                    elif "translate" in query:
                        query=query.replace("mark","")
                        query=query.replace("translate","")   
                        translategl(query)

                    elif "weather" in query or "distance" in query:
                        query=query.replace("mark","")
                        Query(query)

                    elif "write a note" in query or "make a note" in query:
                        writenote()


                    elif "show me the notes" in query or "read notes" in query:
                        speak("Reading Notes")
                        file=open("Notes.txt", "r")
                        data_note =file.readlines()
                        speak(data_note)

                    elif "open" in query:
                        openapp(query)

                    elif "close" in query:
                        closeapp(query)
                        
                    elif "internet speed" in query:
                        Speedtest()

                    elif "pdf reader" in query:
                        Pdfread() 

                    elif "send message" in query:
                        Message()

                    elif "screen recording" in query:
                        Screenrec()

                    elif "how to" in query:
                        search(query)

                    elif "send email" in query:
                        Email()

                    elif "using artificial intelligence" in query:
                        ai(prompt=query)

                    elif "stop listening" in query:
                        speak("Ok sir")
                        break

                    else:
                        #question=takecommand()
                        Reply(query)

startexecution=VirtualassistantMain()

class GuiofVirtualassistant(QWidget):
    def __init__(self):
        super(GuiofVirtualassistant,self).__init__()
        self.gui=Ui_Form()
        self.gui.setupUi(self)

        self.gui.Pushbutton_start.clicked.connect(self.starttask)
        self.gui.Pushbutton_exit.clicked.connect(self.close)


    def starttask(self):
        self.gui.initializing=QtGui.QMovie("Gui/initial.gif")
        self.gui.Gif_1.setMovie(self.gui.initializing)
        self.gui.initializing.start()

        self.gui.loading=QtGui.QMovie("Gui/loading-gif.gif")
        self.gui.Gif_2.setMovie(self.gui.loading)
        self.gui.loading.start()

        self.gui.speaking=QtGui.QMovie("Gui/42787621ed6d40f0c30f0ae423fc572c.gif")
        self.gui.Gif_3.setMovie(self.gui.speaking)
        self.gui.speaking.start()

        self.gui.earth=QtGui.QMovie("Gui/Earth.gif")
        self.gui.Gif_4.setMovie(self.gui.earth)
        self.gui.earth.start()
        
        self.gui.listening=QtGui.QMovie("Gui/6ba174bf48e9b6dc8d8bd19d13c9caa9.gif")
        self.gui.Gif_6.setMovie(self.gui.listening)
        self.gui.listening.start()

        self.gui.loading=QtGui.QMovie("Gui/a7b015d343ad801ad6da8c242dc6ae06.gif")
        self.gui.Gif_7.setMovie(self.gui.loading)
        self.gui.loading.start()

        timer=QTimer(self)
        timer.timeout.connect(self.showTimeLive)
        timer.start(999)

        startexecution.start()

    def showTimeLive(self):
        t_ime=QTime.currentTime()
        time=t_ime.toString()
        d_ate=QDate.currentDate()
        date=d_ate.toString()
        label_time="Time :"+time
        label_date="Date :"+date

        self.gui.Text_time.setText(label_time)
        self.gui.Text_date.setText(label_date)

    def updatemoviesdynamically(self, state):
        if state=="Listening":
            self.gui.Gif_6.raise_()
            self.gui.Gif_3.hide()
            self.gui.Gif_7.hide()
            self.gui.Gif_6.show()
        elif state=="Speaking":
            self.gui.Gif_3.raise_()
            self.gui.Gif_6.hide()
            self.gui.Gif_7.hide()
            self.gui.Gif_3.show()
        elif state=="Loading":
            self.gui.Gif_7.raise_()
            self.gui.Gif_3.hide()
            self.gui.Gif_6.hide()
            self.gui.Gif_7.show()

    def terminalprint(self,text):
        self.gui.textBrowser.append(text)



app=QApplication(sys.argv)
virtualassistant_gui=GuiofVirtualassistant()
virtualassistant_gui.show()
sys.exit(app.exec_())
        
    






