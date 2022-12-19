import pyautogui
import pandas as pd

def readFile():
    x = pyautogui.prompt('Please enter excel timesheet filename', 'LostSG Salary Calculator')
    x = x + ".xlsx"
    print(x)
    df = pd.read_excel(x, sheet_name='Punch Time')
    print(df)
def main():
    readFile()
if __name__ == '__main__':
    main()