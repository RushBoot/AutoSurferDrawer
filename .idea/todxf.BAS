Sub Main
	Dim SurferApp As Object
	Set SurferApp = GetObject(,"Surfer.Application")
	SurferApp.Visible = True
	Set WshShell = CreateObject("WScript.Shell")
	currentPath = WshShell.CurrentDirectory
	folderPath=currentPath+"\.idea\folder_info.xlsx"
Set ExcelApp = CreateObject("Excel.Application")
	ExcelApp.Visible = True
Set ExcelWorkbook = ExcelApp.Workbooks.Open(folderPath)
Set ExcelSheet = ExcelWorkbook.Sheets(1)
	optionpath=ExcelSheet.Range("A6").Text
	Optionpath=optionpath
Set ExcelWorkbook1 = ExcelApp.Workbooks.Open(Optionpath)
Set ExcelSheet1 = ExcelWorkbook1.Sheets(1)
	srf_path=ExcelSheet1.Range("C2").Text
	dxf_path=ExcelSheet1.Range("E2").Text
	data_path=ExcelSheet1.Range("F2").Text
	firstname=ExcelSheet1.Range("I3").Text
	endname=ExcelSheet1.Range("J3").Text
	ExcelFile =data_path
Set ExcelWorkbook = ExcelApp.Workbooks.Open(ExcelFile)
Set ExcelSheet = ExcelWorkbook.Sheets(1)
Dim cell, i
For Each cell In ExcelSheet.Range(firstname+":"+endname)
	i=cell.Text
	If i = "" Then End
	Dim Plot As Object
	Set Plot = SurferApp.Documents.Open(srf_path+"\"+i+".srf")
	Dim Save As Object
	Set Save = SurferApp.Documents(1)
	Plot.Export(dxf_path+"\"+i+".dxf")
    Plot.Close
Next
MsgBox "导出已完成！请关闭该窗口", vbInformation + vbMsgBoxSetForeground, "提示"
ExcelWorkbook.Close False
ExcelApp.Quit
Set ExcelSheet = Nothing
Set ExcelWorkbook = Nothing
Set ExcelApp = Nothing
End Sub
