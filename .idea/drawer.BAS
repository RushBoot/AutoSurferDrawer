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
Dim csv_path,srf_path,grd_path,data_path,basemap_path,blank_path,xmin,ymin,xmax,ymax
	csv_path=ExcelSheet1.Range("A2").Text
	srf_path=ExcelSheet1.Range("C2").Text
	grd_path=ExcelSheet1.Range("D2").Text
	data_path=ExcelSheet1.Range("F2").Text
	basemap_path=ExcelSheet1.Range("G2").Text
	blank_path=ExcelSheet1.Range("H2").Text
	Debug.Print("blank_path:"+blank_path)
	xmin=ExcelSheet1.Range("A3").Text
	ymin=ExcelSheet1.Range("B3").Text
	xmax=ExcelSheet1.Range("C3").Text
	ymax=ExcelSheet1.Range("D3").Text
	x_col=ExcelSheet1.Range("E3").Text
	y_col=ExcelSheet1.Range("F3").Text
	firstname=ExcelSheet1.Range("I3").Text
	endname=ExcelSheet1.Range("J3").Text
	ExcelFile =  data_path' 替换为你的Excel文件路径
Set ExcelWorkbook = ExcelApp.Workbooks.Open(ExcelFile)
Set ExcelSheet = ExcelWorkbook.Sheets(1)
Dim cell, i
For Each cell In ExcelSheet.Range(firstname+":"+endname)
	i=cell.Text
	If i = "" Then End
	Dim Plot As Object
	Set Plot = SurferApp.Documents.Open(basemap_path)
	DataFile = csv_path+"\"+i+".csv"
	OutGrid = grd_path+"\"+i+".grd"
	SurferApp.GridData3 (DataFile:=DataFile, Algorithm:= Surfer.SrfGridAlgorithm.srfKriging, xCol:=x_col, yCol:=y_col, zCol:=3, NumRows:=1000, NumCols:=1000, xmin:=xmin, xmax:=xmax , ymin:=ymin, ymax:=ymax, OutGrid:=Outgrid)
	If blank_path <> "" Then SurferApp.GridBlank (InGrid:=OutGrid, BlankFile:=blank_path, Outgrid:=OutGrid)
	Set MapFrame = Plot.Shapes.AddContourMap(GridFileName:=OutGrid)
	Dim ContourLayer As Object
	Set ContourLayer = MapFrame.Overlays(1)
	ContourLayer.LevelMajorInterval = 1
	ContourLayer.MajorLine.ForeColorRGBA.Color = srfColorBlue
	ContourLayer.FillContours = True
	ContourLayer.FillForegroundColorMap.LoadPreset("Geology2")
	ContourLayer.ShowColorScale=True
	ContourLayer.LabelFont.ForeColorRGBA.Color = srfColorBlue
	ContourLayer.LabelFont.Face = "Times New Roman"
	ContourLayer.ColorScale.LabelFont.Face = "Times New Roman"
	Dim a,b,c,d
	a=ContourLayer.ColorScale.Height
	b=ContourLayer.ColorScale.Width
	c=3*b/a
	ContourLayer.ColorScale.Height = 3
	ContourLayer.ColorScale.Width = c
	ContourLayer.ColorScale.Left = 7
	ContourLayer.ColorScale.Top = 7
	MapFrame.SetZOrder (Zorder:=srfZOToBack)
	Dim Save As Object
	Set Save = SurferApp.Documents(1)
	Plot.SaveAs srf_path+"\"+i+".srf"
Next
MsgBox "绘图已完成！请关闭该窗口", vbInformation + vbMsgBoxSetForeground, "提示"
ExcelWorkbook1.Save
ExcelWorkbook1.Close True
ExcelWorkbook.Save
ExcelWorkbook.Close True
ExcelApp.Quit
Set ExcelSheet = Nothing
Set ExcelWorkbook = Nothing
Set ExcelApp = Nothing


End Sub
