import speech_recognition as sr

def voice_assitant():
    r = sr.Recognizer()
    mic = sr.Microphone()
    with mic as source:
        r.adjust_for_ambient_noise(source)
        audio = r.listen(source)
        try:
            transcript = r.recognize_google(audio,language = 'en-IN')
        except:
            transcript = 'Api Error'
        return transcript

if __name__  == "__main__":
    print(voice_assitant())
