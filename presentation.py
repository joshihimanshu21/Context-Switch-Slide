import win32com.client
import time
from try_mic import voice_assitant

app = win32com.client.Dispatch("PowerPoint.Application")

presentation = app.Presentations.Open(FileName=u"C:\\Users\\ejimhos\\Downloads\\CV Presentation.pptx")


if __name__ == "__main__":
    presentation.SlideShowSettings.Run()
    while(1):
        command = voice_assitant()
        print("You said {}".format(command))
        if command.lower() == 'next':
            presentation.SlideShowWindow.View.Next()
        elif command.lower() == 'close':
            presentation.SlideShowWindow.View.Exit()
            break
    app.Quit()
