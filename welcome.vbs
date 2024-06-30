Set xlsxApp=CreateObject("Excel.Application")
xlsxApp.Visible=true

Set xlsxWorkbook= xlsxApp.Workbooks.Open("C:\Users\Nitish Dhawan\OneDrive - BENNETT UNIVERSITY\Desktop\Destination2.xlsm")
xlsxApp.Run("MyWelcomeMessage")


