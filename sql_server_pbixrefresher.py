import os
import subprocess
import time
import pyautogui


class PbiRefresher:

    def open_file(self):
        pyautogui.FAILSAFE = False
        print("Opening the pbix file . . .")
        intro_path = os.path.dirname(__file__)
        button_file = r"/file_icon.png"
        full_path = "".join(f"{intro_path}{button_file}")
        button_fileopen = pyautogui.locateOnScreen(full_path, confidence=0.9)
        while button_fileopen is None:
            time.sleep(2)
            button_fileopen = pyautogui.locateOnScreen(full_path, confidence=0.9)
        buttonx, buttony = pyautogui.center(button_fileopen)
        pyautogui.click(buttonx, buttony)

        time.sleep(1)
        pyautogui.press('down')
        time.sleep(1)
        pyautogui.press('down')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1)

        button_zeppa = r"/zeppa_icon.png"
        full_path = "".join(f"{intro_path}{button_zeppa}")
        button_zeppafile = pyautogui.locateOnScreen(full_path, confidence=0.9)
        while button_zeppafile is None:
            time.sleep(2)
            button_zeppafile = pyautogui.locateOnScreen(full_path, confidence=0.9)
        buttonx, buttony = pyautogui.center(button_zeppafile)
        pyautogui.click(buttonx, buttony)
        time.sleep(1)
        pyautogui.press('enter')

        button_date_spec = r"/ampop_icon.png"
        full_path = "".join(f"{intro_path}{button_date_spec}")
        button_date = pyautogui.locateOnScreen(full_path, confidence=0.9)
        while button_date is None:
            time.sleep(2)
            button_date = pyautogui.locateOnScreen(full_path, confidence=0.9)
            
        time.sleep(7)
        pyautogui.hotkey('alt', 'h')
        time.sleep(1)
        pyautogui.typewrite('r', interval=0)
        time.sleep(10)
        pyautogui.hotkey('alt', 'h')
        time.sleep(1)
        pyautogui.typewrite('r', interval=0)
        print("Refresh Done!")
        self.saving_file()

    def saving_file(self):
        # save
        time.sleep(8)
        pyautogui.hotkey('ctrl', 's')
        time.sleep(1)
        pyautogui.hotkey('ctrl', 's')
        print("File was saved !")
        self.publish_file()

    def publish_file(self):
        # publish
        time.sleep(5)
        pyautogui.hotkey('alt', 'h')
        time.sleep(1)
        pyautogui.typewrite('pu', interval=0)
        time.sleep(3)
        pyautogui.press('enter')
        time.sleep(3)
        pyautogui.press('enter')
        print("Published !")
        self.closing_app()

    def closing_app(self):
        # closing
        intro_path = os.path.dirname(__file__)
        button_success = r"/success_icon.png"
        full_path = "".join(f"{intro_path}{button_success}")
        button_finished = pyautogui.locateOnScreen(full_path, confidence=0.9)
        while button_finished is None:
            time.sleep(2)
            button_finished = pyautogui.locateOnScreen(full_path, confidence=0.9)

        button_got = r"/got_it_icon.png"
        full_path = "".join(f"{intro_path}{button_got}")
        button_gotit = pyautogui.locateOnScreen(full_path, confidence=0.9)
        while button_gotit is None:
            time.sleep(2)
            button_gotit = pyautogui.locateOnScreen(full_path, confidence=0.9)
        buttonx, buttony = pyautogui.center(button_gotit)
        pyautogui.click(buttonx, buttony)
        print("Sub window closed !")

        # close the program
        time.sleep(2)
        pyautogui.hotkey('ctrl', 's')
        time.sleep(5)
        pyautogui.hotkey('alt', 'f4')
        time.sleep(2)
        pyautogui.press('enter')
        print("Program was closed ! \n PBI ZEPPA Disbursements file was successfully refreshed, saved and published !")


if __name__ == "__main__":

    path_pbi_file = r"C:\Program Files\Microsoft Power BI Desktop\bin\PBIDesktop.exe"
    p = subprocess.Popen(path_pbi_file, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                         universal_newlines=True, shell=False)
    time.sleep(18)
    print("The process started . . .")
    pyautogui.FAILSAFE = False
    # pbix = PbiRefresher()
    # pbix.open_file()
