from win32com.client.gencache import EnsureDispatch
from win32com.client import constants

xl = EnsureDispatch('Excel.Application')
xl.Workbooks.Open('C:\\Users\\Ozgur\\Desktop\\htm\\macro2.xlsm')
xl.DisplayAlerts=False
xl.Application.Run("macro2.xlsm!module1.xx")
xl.Application.Save()
xl.Application.Quit()
del xl
