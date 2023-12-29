# *************************************************************************************
# *   MIT License                                                                     *
# *                                                                                   *
# *   Copyright (c) 2023 Paul Ebbers                                                  *
# *                                                                                   *
# *   Permission is hereby granted, free of charge, to any person obtaining a copy    *
# *   of this software and associated documentation files (the "Software"), to deal   *
# *   in the Software without restriction, including without limitation the rights    *
# *   to use, copy, modify, merge, publish, distribute, sublicense, and/or sell       *
# *   copies of the Software, and to permit persons to whom the Software is           *
# *   furnished to do so, subject to the following conditions:                        *
# *                                                                                   *
# *   The above copyright notice and this permission notice shall be included in all  *
# *   copies or substantial portions of the Software.                                 *
# *                                                                                   *
# *   THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR      *
# *   IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,        *
# *   FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE     *
# *   AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER          *
# *   LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,   *
# *   OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE   *
# *   SOFTWARE.                                                                       *
# *                                                                                   *
# *************************************************************************************/

class StandardFunctions_FreeCAD:
    def Mbox(text, title="", style=0, default="", stringList="[,]"):
   """
    Message Styles:\n
    0 : OK                          (text, title, style)\n
    1 : Yes | No                    (text, title, style)\n
    20 : Inputbox                    (text, title, style, default)\n
    21 : Inputbox with dropdown      (text, title, style, default, stringlist)\n
    """
    from PySide2.QtWidgets import QMessageBox, QInputDialog

    Icon = QMessageBox.Information
    if IconType == "NoIcon":
        Icon = QMessageBox.NoIcon
    if IconType == "Question":
        Icon = QMessageBox.Question
    if IconType == "Warning":
        Icon = QMessageBox.Warning
    if IconType == "Critical":
        Icon = QMessageBox.Critical

    if style == 0:
        # Set the messagebox
        msgBox = QMessageBox()
        msgBox.setIcon(Icon)
        msgBox.setText(text)
        msgBox.setWindowTitle(title)

        reply = msgBox.exec_()
        return reply
    if style == 1:
        # Set the messagebox
        msgBox = QMessageBox()
        msgBox.setIcon(Icon)
        msgBox.setText(text)
        msgBox.setWindowTitle(title)
        # Set the buttons and default button
        msgBox.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        msgBox.setDefaultButton(QMessageBox.No)

        reply = msgBox.exec_()
        if reply == QMessageBox.Yes:
            return "yes"
        if reply == QMessageBox.No:
            return "no"
    if style == 20:
        reply = QInputDialog.getText(parent=None, title=title, label=text, text=default)
        if reply[1]:
            # user clicked OK
            replyText = reply[0]
        else:
            # user clicked Cancel
            replyText = reply[0]  # which will be "" if they clicked Cancel
        return str(replyText)
    if style == 21:
        reply = QInputDialog.getItem(parent=None, title=title, label=text, items=stringList, current=1, editable=True)
        if reply[1]:
            # user clicked OK
            replyText = reply[0]
        else:
            # user clicked Cancel
            replyText = reply[0]  # which will be "" if they clicked Cancel
        return str(replyText)

    def SaveDialog(files, OverWrite: bool = True):
        """
        files must be like:\n
        files = [\n
            ('All Files', '*.*'),\n
            ('Python Files', '*.py'),\n
            ('Text Document', '*.txt')\n
        ]\n
        \n
        OverWrite:\n
        If True, file will be overwritten\n
        If False, only the path+filename will be returned\n
        """
        import tkinter as tk
        from tkinter.filedialog import asksaveasfile
        from tkinter.filedialog import askopenfilename

        # Create the window
        root = tk.Tk()
        # Hide the window
        root.withdraw()

        if OverWrite is True:
            file = asksaveasfile(filetypes=files, defaultextension=files)
            if file is not None:
                return file.name
        if OverWrite is False:
            file = askopenfilename(filetypes=files, defaultextension=files)
            if file is not None:
                return file

    def GetLetterFromNumber(number: int, UCase: bool = True):
        from openpyxl.utils import get_column_letter

        Letter = get_column_letter(number)

        # If UCase is true, convert to upper case
        if UCase is True:
            Letter = Letter.upper()

        return Letter

    def GetNumberFromLetter(Letter):
        from openpyxl.utils import column_index_from_string

        Number = column_index_from_string(Letter)

        return Number

    def GetA1fromR1C1(input: str) -> str:
        if input[:1] == "'":
            input = input[1:]
        try:
            input = input.upper()
            ColumnPosition = input.find("C")
            RowNumber = int(input[1:(ColumnPosition)])
            ColumnNumber = int(input[(ColumnPosition + 1):])

            ColumnLetter = GetLetterFromNumber(ColumnNumber)

            return str(ColumnLetter + str(RowNumber))
        except Exception:
            return ""

    def CheckIfWorkbookExists(FullFileName: str, CreateIfNone: bool = True):
        import os
        from openpyxl import Workbook

        result = False
        try:
            result = os.path.exists(FullFileName)
        except Exception:
            if CreateIfNone is True:
                Filter = [
                    ("Excel", "*.xlsx"),
                    (
                        "Excel Macro-enabled Workbook",
                        "*.xlsm",
                    ),
                ]
                FullFileName = SaveDialog(Filter)
                if FullFileName.strip():
                    wb = Workbook(str(FullFileName))
                    wb.save(FullFileName)
                    wb.close()
                    result = True
            if CreateIfNone is False:
                result = False
        return result

    def ColorConvertor(ColorRGB: [], Alpha: float = 1) -> ():
        """
        A single function to convert colors to rgba colors as a tuple of float from 0-1
        ColorRGB:   [255,255,255]
        Alpha:      0-1
        """
        from matplotlib import colors as mcolors

        ColorRed = ColorRGB[0] / 255
        colorGreen = ColorRGB[1] / 255
        colorBlue = ColorRGB[2] / 255

        color = (ColorRed, colorGreen, colorBlue)

        result = mcolors.to_rgba(color, Alpha)

        return result

    def OpenFile(FileName: str):
        """
        Filename = full path with filename as string
        """
        import subprocess
        import os
        import platform

        if os.path.exists(FileName):
            if platform.system() == "Darwin":  # macOS
                subprocess.call(("open", FileName))
            elif platform.system() == "Windows":  # Windows
                os.startfile(FileName)
            else:  # linux variants
                subprocess.call(("xdg-open", FileName))
        else:
            print(f"Error: {FileName} does not exist.")

    def SetColumnWidth_SpreadSheet(self, sheet, column: str, cellValue: str, factor: int = 10) -> bool:
        """_summary_

        Args:
            sheet (_type_): FreeCAD spreadsheet object.\n
            column (str): The column for which the width will be set. must be like "A", "B", etc.\n
            cellValue (str): The string to calulate the widht from.\n
            factor (int, optional): to increase the stringlength with a factor. Defaults to 10.\n

        Returns:
            bool: returns True or False
        """
        try:
            # Calculate the text length needed.
            length = int(len(cellValue) * factor)

            print(column)
            # Set the column width
            sheet.setColumnWidth(column, length)

            # Recompute the sheet
            sheet.recompute()
        except Exception:
            return False

        return True

def Print(Input: str, Type: str = ""):
    """_summary_

    Args:
        Input (str): Text to print.\n
        Type (str, optional): Type of message. (enter Warning, Error or Log). Defaults to "".
    """
    import FreeCAD as App

    if Type == "Warning":
        App.Console.PrintWarning(Input + "\n")
    elif Type == "Error":
        App.Console.PrintError(Input + "\n")
    elif Type == "Log":
        App.Console.PrintLog(Input + "\n")
    else:
        App.Console.PrintMessage(Input + "\n")
