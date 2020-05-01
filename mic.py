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
            return ''
        return transcript

def autorecognize():
    while(1):
        text = voice_assitant()
        if text=='stop':
            break
        elif text == '':
            continue
        else:
            print(text)

if __name__=="__main__":
    autorecognize()
