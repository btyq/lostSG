import pyautogui
import openpyxl

class Employee:
    def __init__(self, name):
        self.name = name
        
def readFile():
    x = "temp"
    try:
        x = pyautogui.prompt('Please enter excel timesheet filename', 'LostSG Salary Calculator')
        if x is None:
            pass
        else:
            x = x + ".xlsx"
            print(x)
            punchTime = openpyxl.load_workbook(x)
            punchTime.active = punchTime['Punch Time']
            sh = punchTime.active
            for i in range(1, sh.max_row+1):
                print("\n")
                print("Row ", i, " data :")
                for j in range(1, sh.max_column+1):
                    cell_obj = sh.cell(row=i, column=j)
                    print(cell_obj.value, end=" ")

    except FileNotFoundError:
        pyautogui.alert(text='File not Found', title='Error', button='OK')

def main():
    readFile()
if __name__ == '__main__':
    main()