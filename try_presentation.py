import win32com.client
import time
from try_mic import voice_assitant

app = win32com.client.Dispatch("PowerPoint.Application")

presentation = app.Presentations.Open(FileName=u"C:\\Users\\ejimhos\\Downloads\\CV Presentation.pptx",ReadOnly = 1)


if __name__ == "__main__":
    presentation.SlideShowSettings.Run()
    c = int(input('Enter the slide to which we need to skip: '))
    while(c):
    	presentation.SlideShowWindow.View.Next()
    	c-=1
    time.sleep(2)
    presentation.SlideShowWindow.View.Exit()
    app.Quit()
