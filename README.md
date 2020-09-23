<div align="center">

## Server Side PDF


</div>

### Description

Generate PDF files on the server without any server-side components. Based on X2PDF.NET library created by Arne Garvander. Limitations: Paragraph (TextArea) cannot exceed on page. Table cannot exceed one page. To read about PDF file specifications go to: http://partners.adobe.com/asn/tech/pdf/specifications.jsp

To learn how to edit an existing PDF file go to: http://www.15seconds.com/issue/990902.htm

To learn how to merge PDF file go to: http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=37121&lngWId=1
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Igor Krupitsky](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/igor-krupitsky.md)
**Level**          |Intermediate
**User Rating**    |5.0 (30 globes from 6 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Files](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files__4-2.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/igor-krupitsky-server-side-pdf__4-8633/archive/master.zip)





### Source Code

```
<%@ Language=VBScript %>
<%
Option Explicit
Response.Expires = 0
Public Const Fonts_Helvetica = 0
Public Const Fonts_Courier = 1
Public Const Fonts_Times_Roman = 2
Public Const FontStyles_Regular = 0
Public Const FontStyles_Bold = 1
Public Const FontStyles_Italic = 2
Public Const FontStyles_BoldItalic = 3
Public Const Borders_thick = 1
Public Const Borders_thin = 2
Public Const Borders_none = 3
'===================
Dim oPdf 'As PDFDocument
Dim sText 'As String
Dim oTexts 'As TextArea
Dim oTable 'As table
Dim oRow 'As row
Dim oCell 'As cell
Set oPdf = New PDFDocument
oPdf.Creator = "Igor Krupitsky"
Set oTexts = New TextArea
oTexts.AddText "Server side PDF rules!", Fonts_Times_Roman, 15, ""
oTexts.AddText "Planet Source Code.", Fonts_Courier, 15, FontStyles_Bold
oTexts.AddText "The largest Public source code database on the Internet With 8,297,283 lines of code, articles and tutorials in 11 languages,as well as 1,127 open job postings.", Fonts_Courier, 12, ""
oPdf.AddControl oTexts
Set oTable = New Table
oTable.Border = Borders_thin 'Borders_none, Borders_thick
Set oRow = New row
Set oCell = New cell
oCell.AddText "First Name", Fonts_Helvetica, 10
oRow.AddCell oCell
Set oCell = New cell
oCell.AddText "Last Name", Fonts_Helvetica, 10
oRow.AddCell oCell
Set oCell = New cell
oCell.AddText "Phone", Fonts_Helvetica, 10
oRow.AddCell oCell
oTable.AddRow oRow
Set oRow = New row
Set oCell = New cell
oCell.AddText "James", Fonts_Helvetica, 14
oRow.AddCell oCell
Set oCell = New cell
oCell.AddText "Bond", Fonts_Helvetica, 14
oRow.AddCell oCell
Set oCell = New cell
oCell.AddText "007", Fonts_Helvetica, 14
oRow.AddCell oCell
oTable.AddRow oRow
oPdf.AddControl oTable
'oPdf.OutputToFile "c:\temp\test.pdf"
Dim sTemp: sTemp = oPdf.OutputToStream()
Response.ContentType = "application/pdf"
Response.BinaryWrite StringToMultiByte(sTemp)
'===================
Class Cell
	Public default Property Get ClassName() 'As FontStyles
		ClassName = "Cell"
	End Property
	Private m_textArea 'As TextArea
	Private m_Height 'As Integer ' PDFUnits
	Public ColumnSpan 'As Integer
	Public WidthInPDFUnits 'As Integer
	Public StartPDFH 'As Integer ' Start of text
	Public StartPDFV 'As Integer
	Public WidthInPercent 'As Integer
	Private Sub Class_Initialize()
	 Set m_textArea = New TextArea
	 ColumnSpan = 1
	End Sub
	function GetCopy() 'As cell
	 Dim myCell 'As cell
	 Dim myText 'As TextObject
	 Set myCell = New cell
	 With myCell
	 For Each myText In m_textArea.getTexts
	.AddText myText.Text, myText.Font, myText.FontSize
	 Next
	 .ColumnSpan = ColumnSpan
	 End With
	 Set GetCopy = myCell
	End function
	function Draw(ByRef FontAlias, ByRef pagenum, ByVal TopMargin) 'As PDFObject
	 m_textArea.StartPDFH = StartPDFH
	 Set Draw = m_textArea.Draw(StartPDFV, WidthInPDFUnits, FontAlias, pagenum, TopMargin)
	End function
	Public Sub AddText(ByVal Text, ByVal Font, ByVal FontSize)
	 if Font = "" Then Font = Fonts_Helvetica
	 if FontSize = "" Then FontSize = 10
	 m_textArea.AddText Text, Font, FontSize, FontStyles_Regular
	End Sub
	function CalculateHeight(ByVal width) 'As Integer
	 WidthInPDFUnits = width
	 m_textArea.CalculateHeight (width)
	 m_Height = m_textArea.HeightInPDFunits
	 CalculateHeight = m_Height
	End function
End Class
'===================
Class CFontObj
	Public default Property Get ClassName() 'As FontStyles
		ClassName = "FontObj"
	End Property
	Dim m_Font 'As Fonts
	Dim m_FontName 'As String
	Dim m_fontStyle 'As FontStyles
	Public FontRef 'As String
	Public FontObj 'As String
	Private Sub Class_Initialize()
	 m_Font = Fonts_Helvetica
	 m_fontStyle = FontStyles_Regular
	 m_FontName = ""
	End Sub
	function equals(ByVal FontObj) 'As Boolean
	 equals = True
	 if m_Font <> FontObj.Font Or m_fontStyle <> FontObj.FontStyle Then
	 equals = False
	 Else
	 equals = True
	 End if
	End function
	Public Property Get FontStyle() 'As FontStyles
	FontStyle = m_fontStyle
	End Property
	Public Property Let FontStyle(ByVal myFontStyle)
	m_fontStyle = myFontStyle
	Call SetFontName
	End Property
	Public function ValidFont(ByVal Font) 'As Boolean
	if -1 < Font And Font < 5 Then
	 ValidFont = True
	Else
	 ValidFont = False
	End if
	End function
	Public Property Get HorizontalSpace() 'As Double
	 Dim space 'As Double
	 Select Case m_Font
	 Case Fonts_Courier
	space = 1.7
	 Case Fonts_Helvetica
	space = 2.2
	 Case Fonts_Times_Roman
	space = 2.4
	 Case Else
	space = 2
	 End Select
	 if m_fontStyle = FontStyles_Bold Or m_fontStyle = FontStyles_BoldItalic Then
	 space = space * 0.91
	 End if
	 HorizontalSpace = space
	End Property
	Public Property Get Font() 'As Fonts
	 Font = m_Font
	End Property
	Public Property Let Font(ByVal myFont)
	 m_Font = myFont
	 Call SetFontName
	End Property
	Private Sub SetFontName()
	 Select Case m_Font
	 Case Fonts_Courier
	Select Case m_fontStyle
	Case FontStyles_Regular
	 m_FontName = "Courier"
	Case FontStyles_Bold
	 m_FontName = "Courier-Bold"
	Case FontStyles_Italic
	 m_FontName = "Courier-Oblique"
	Case FontStyles_BoldItalic
	 m_FontName = "Courier-BoldOblique"
	Case Else
	 Err.Raise 100,"","Invalid Font style."
	End Select
	 Case Fonts_Helvetica
	Select Case m_fontStyle
	Case FontStyles_Regular
	 m_FontName = "Helvetica"
	Case FontStyles_Bold
	 m_FontName = "Helvetica-Bold"
	Case FontStyles_Italic
	 m_FontName = "Helvetica-Oblique"
	Case FontStyles_BoldItalic
	 m_FontName = "Helvetica-BoldOblique"
	Case Else
	 Err.Raise 100,"","Invalid Font style."
	End Select
	 Case Fonts_Times_Roman
	Select Case m_fontStyle
	Case FontStyles_Regular
	 m_FontName = "Times-Roman"
	Case FontStyles_Bold
	 m_FontName = "Times-Bold"
	Case FontStyles_Italic
	 m_FontName = "Times-Italic"
	Case FontStyles_BoldItalic
	 m_FontName = "Times-BoldItalic"
	Case Else
	 Err.Raise 100,"","Invalid Font style."
	End Select
	 Case Else
	Err.Raise 100,"","Invalid Font"
	 End Select
	End Sub
	Public Property Get FontName() 'As String
	 FontName = m_FontName
	End Property
End Class
'===================
Class PageBreak
	Public default Property Get ClassName() 'As FontStyles
		ClassName = "PageBreak"
	End Property
	Public StartPDFH 'As Integer
	function GetCopy()
	 Dim pg 'As PageBreak
	 pg = New PageBreak
	 pg.StartPDFH = StartPDFH
	 GetCopy = pg
	End function
	function Draw(ByRef StartV, ByVal width, ByRef FontAlias, _
	 ByRef pagenum, ByVal TopStart) 'As PDFObject
	 Dim stream 'As String
	 Dim PDFO 'As PDFObject
	 Dim mid 'As Integer
	 Set PDFO = New PDFObject
	 mid = StartPDFH + width / 2
	 pagenum = pagenum + 1
	 stream = "BT" & vbCr
	 stream = stream & "/F1 " & 10 & " Tf" & vbCr
	 stream = stream & "1 0 0 1 " & mid & " 50 Tm" & vbCr
	 stream = stream & "(" & pagenum & ") Tj" & vbCr
	 stream = stream & "/F1 " & 6 & " Tf" & vbCr
	 stream = stream & "1 0 0 1 " & StartPDFH & " 50 Tm" & vbCr
	 stream = stream & "(Copyright Arne@garvander.com) Tj" & vbCr
	 stream = stream & "ET" & vbCr
	 PDFO.addStream (stream)
	 Set Draw = PDFO
	End function
	function toString() 'As String
	 toString = "Page Break"
	End function
End Class
'===================
Class PDFDocument
	Public default Property Get ClassName() 'As FontStyles
		ClassName = "PDFDocument"
	End Property
	Dim m_Title 'As String
	Dim m_keywords 'As String
	Dim m_subject 'As String
	Dim m_FontAlias 'As Scripting.Dictionary ' One entry per font
	Dim m_PageNumber 'As Integer
	Public Author 'As String
	Public Creator 'As String
	Public Producer 'As String
	Public OutputFileName 'As String
	Dim m_OutputStream
	Dim m_OutputToStream 'As Boolean
	Dim Position 'As Integer
	Dim m_PDFLocation(5000) 'As Integer ' Positions of all the PDF objects
	Dim pageObj(5000) 'As Integer' Page objects
	Dim obj 'As Integer ' PDF objects
	Dim m_rootObj 'As Integer' RootObject is the object after properties
	Dim m_TopPagesObj 'As Integer ' Top page comes after rootobject
	Dim m_EncodingObj 'As Integer ' Object For Encoding Type
	Dim m_PropObj 'As Integer
	Dim cache 'As String
	Dim m_controls 'As Scripting.Dictionary
	Dim m_PageHeight 'As Integer
	Dim m_Pagewidth 'As Integer
	Dim m_drawableWidth 'As Integer
	Dim m_TopMargin 'As Integer ' 3/4 inch, An adobe document has another 1/4 inch built in margin
	Dim m_LeftMargin 'As Integer ' 1 inch, An adobe document has another 1/4 inch built in margin
	Private Sub Class_Initialize()
	 m_Pagewidth = 612
	 m_PageHeight = 792
	 m_TopMargin = 54
	 m_LeftMargin = 72
	 Set m_controls = CreateObject("Scripting.Dictionary")
	 Set m_FontAlias = CreateObject("Scripting.Dictionary")
	 obj = 0
	 Position = 0
	 cache = ""
	 m_OutputToStream = False
	End Sub
	Public Property Get PageWidth() 'As Integer
	 PageWidth = m_Pagewidth / 72
	End Property
	Public Sub AddControl(ByVal control)
	 Dim ta 'As TextArea
	 if TypeName(control) = "TextArea" Then
	 Set ta = control.GetCopy
	 m_controls.Add ta, ""
	 Else
	 m_controls.Add control, ""
	 End if
	End Sub
	Public Sub OutputToFile(ByVal filename)
	 if filename <> "" Then
	 OutputFileName = filename
	 End if
	 if FileExists(OutputFileName) Then
	 Kill (OutputFileName)
	 End if
	 Call WriteStart
	 Call WriteHead
	 Call WritePage
	 Call endPDF
	End Sub
	Public function OutputToStream()
		m_OutputToStream = True
	 Call WriteStart
	 Call WriteHead
	 Call WritePage
	 Call endPDF
	 OutputToStream = m_OutputStream
	 m_OutputToStream = False
	End function
	Private function WritePage()
	 Dim beginstream 'As String
	 Dim Fonts 'As String
	 Dim FontRef
	 Dim key 'As String
	 Dim PDFO 'As PDFObject
	 Dim fonto 'As FontObj
	 Dim Resources 'As String
	 Dim contents 'As String
	 Dim stream 'As String
	 Dim StartY 'As Integer
	 Dim width 'As Integer
	 Dim control
	 Dim dummy 'As String
	 Dim page 'As PageBreak
	 Dim PageFonts 'As PDFObject
	 Dim TopStart 'As Integer
	 Set PageFonts = New PDFObject
	 Fonts = " /Font << "
	 StartY = m_PageHeight - m_TopMargin
	 TopStart = StartY
	 width = m_Pagewidth - 2 * m_LeftMargin
	 For Each control In m_controls
	 dummy = control.toString' Debug statement
	 if control.StartPDFH = 0 Then
	control.StartPDFH = m_LeftMargin
	 End if
	 Set PDFO = control.Draw(StartY, width, m_FontAlias, m_PageNumber, TopStart)
	 if PDFO.count > 1 Then
	stream = stream + PDFO.getStream()
	StartPage contents, Resources, stream, Fonts
	stream = ""
	Set PageFonts = New PDFObject
	Fonts = " /Font << "
	 End if
	 stream = stream + PDFO.getStream()
	 Call WriteNewFonts
	 For Each FontRef In PDFO.m_fonts
	Set fonto = m_FontAlias.Item(FontRef)
	if PageFonts.FontExists(fonto.FontObj) = False Then
	if Not PageFonts.m_fonts.Exists(fonto.FontObj) Then
	 PageFonts.m_fonts.Add fonto.FontObj, ""
	End if
	Fonts = Fonts + "/F" & FontRef & fonto.FontObj & " 0 R "
	End if
	 Next
	 Next
	 if Len(stream) Then
	 Set page = New PageBreak
	 page.StartPDFH = m_LeftMargin
	 Set PDFO = page.Draw(StartY, width, m_FontAlias, m_PageNumber, TopStart)
	 stream = stream + PDFO.getStream()
	 StartPage contents, Resources, stream, Fonts
	 End if
	End function
	Private Sub StartPage(ByVal contents, ByVal Resources, ByVal stream, ByVal Fonts)
	 Fonts = Fonts + ">>"
	 Resources = Resources + Fonts + vbCrLf
	 Resources = Resources + "/Procset [/PDF /Text]"
	 obj = obj + 1
	 contents = contents + CStr(obj) & " 0 R"
	 m_PDFLocation(obj) = Position
	 writepdf obj & " 0 obj", False
	 writepdf "<< /Length " & Len(stream) & " >>", False
	 writepdf "stream", False
	 writepdf stream, False
	 writepdf "endstream", False
	 writepdf "endobj", False
	 obj = obj + 1
	 m_PDFLocation(obj) = Position
	 pageObj(m_PageNumber) = obj
	 writepdf obj & " 0 obj", False
	 writepdf "<<", False
	 writepdf "/Type /Page", False
	 writepdf "/Parent " & m_TopPagesObj & " 0 R", False
	 writepdf "/Resources << " & Resources & " >> ", False
	 writepdf "/Contents " & contents, False
	 writepdf ">>", False
	 writepdf "endobj", False
	End Sub
	Private Sub WriteNewFonts()
	 Dim i 'As Integer
	 Dim Fonts 'As String
	 Dim key 'As String
	 Dim fonto 'As FontObj
	 Dim FontName 'As String
	 Dim fontNumber 'As Integer
	 Dim sobj 'As Integer
	 sobj = obj
	 For i = 1 To m_FontAlias.count
	 key = Trim(CStr(i))
	 Set fonto = m_FontAlias.Item(key)
	 if fonto.FontObj = "" Then
	obj = obj + 1
	fonto.FontObj = " " & CStr(obj)
	m_PDFLocation(obj) = Position
	writepdf obj & " 0 obj", False
	writepdf "<<", False
	writepdf "/Type /Font", False
	writepdf "/Subtype /Type1", False ' Adobe Type 1
	writepdf "/Name /F" & fonto.FontRef, False
	writepdf "/BaseEncoding /WinAnsiEncoding", False
	writepdf "/BaseFont /" & fonto.FontName, False
	writepdf ">>", False
	writepdf "endobj", False
	 End if
	 Next
	End Sub
	Private Sub WriteHead()
	 WriteProperties
	 obj = obj + 1
	 m_rootObj = obj ' The root object will be written at the End
	 obj = obj + 1
	 m_TopPagesObj = obj' The Pages object will be written at the End
	 obj = obj + 1
	End Sub
	Private Sub writepdf(ByRef stre, ByRef flush)
	 if flush = "" Then flush = False
		if m_OutputToStream = True Then
			m_OutputStream = m_OutputStream & stre & vbCrLf
			Exit Sub
		End if
	 ' On Error Resume Next
	 Dim i 'As Integer
	 Dim fso 'As FileSystemObject
	 Dim oFile 'As Scripting.TextStream
	 Set fso = CreateObject("Scripting.FileSystemObject")
	 Position = Position + Len(stre) ' Position For the Next object
	 cache = cache & stre & vbCrLf
	 if Len(cache) > 32000 Or flush Then
	 Set oFile = fso.OpenTextFile(OutputFileName, 8, True)
	 oFile.Write cache
	 oFile.Close
	 cache = ""
	 End if
	End Sub
	Private Sub WriteStart()
	 writepdf "%PDF-1.2", False ' Acrobat version 3.0
	 writepdf "%âãÏÓ", False
	End Sub
	Sub endPDF()
	 Dim ty 'As String
	 Dim i 'As Integer
	 Dim xreF 'As Integer
	 m_PDFLocation(m_rootObj) = Position
	 writepdf m_rootObj & " 0 obj", False
	 writepdf "<<", False
	 writepdf "/Type /Catalog", False
	 writepdf "/Pages " & m_TopPagesObj & " 0 R", False
	 writepdf ">>", False
	 writepdf "endobj", False
	 m_PDFLocation(m_TopPagesObj) = Position
	 writepdf m_TopPagesObj & " 0 obj", False
	 writepdf "<<", False
	 writepdf "/Type /Pages", False
	 writepdf "/Count " & m_PageNumber, False
	 writepdf "/MediaBox [ 0 0 " & m_Pagewidth & " " & m_PageHeight & " ]", False
	 ty = ("/Kids [ ")
	 For i = 1 To m_PageNumber
	 ty = ty & pageObj(i) & " 0 R "
	 Next
	 ty = ty & "]"
	 writepdf ty, False
	 writepdf ">>", False
	 writepdf "endobj", False
	 ' Xref
	 xreF = Position
	 writepdf "0 " & obj + 1, False
	 writepdf "0000000000 65535 f ", ""
	 For i = 1 To obj
	 writepdf Right("0000000000" & m_PDFLocation(i), 10) & " 00000 n", False
	 Next
	 ' Trailer
	 writepdf "trailer", False
	 writepdf "<<", False
	 writepdf "/Size " & obj + 1, False
	 writepdf "/Root " & m_rootObj & " 0 R", False
	 writepdf "/Info " & m_PropObj & " 0 R", False
	 writepdf ">>", False
	 writepdf "startxref", False
	 writepdf CStr(xreF), False
	 writepdf "%%EOF", True
	End Sub
	Private Sub WriteProperties()
	 Dim CreationDate 'As String
	 CreationDate = "D:" & GetPdfFormatedDate()
	 obj = obj + 1
	 m_PDFLocation(obj) = Position
	 m_PropObj = obj
	 writepdf obj & " 0 obj", False
	 writepdf "<<", False
	 writepdf "/Author (" & Author & ")", False
	 writepdf "/CreationDate (" & CreationDate & ")", False
	 writepdf "/Creator (" & Creator & ")", False
	 writepdf "/Producer (" & Producer & ")", False
	 writepdf "/Title (" & m_Title & ")", False
	 writepdf "/Subject (" & m_subject & ")", False
	 writepdf "/Keywords (" & m_keywords & ")", False
	 writepdf ">>", False
	 writepdf "endobj", False
	End Sub
	Public function FileExists(ByVal filename) 'As Boolean
	 On Error Resume Next
	 FileExists = FileLen(filename) > 0
	 Err.Clear
	End function
End Class
'===================
Class PDFObject
	Public default Property Get ClassName() 'As FontStyles
		ClassName = "PDFObject"
	End Property
	Dim m_resources 'As String
	Public m_fonts 'As Scripting.Dictionary
	Private m_streams 'As Scripting.Dictionary
	Public PageBreak 'As Boolean
	Private Sub Class_Initialize()
	  Set m_fonts = CreateObject("Scripting.Dictionary")
	  Set m_streams = CreateObject("Scripting.Dictionary")
	End Sub
	Public Sub addStream(ByVal stream)
	  m_streams.Add stream, ""
	End Sub
	Public Function FontExists(ByVal Font) 'As Boolean
	  Dim FontObj 'As String
	  ' FontExists = False
	  For Each FontObj In m_fonts
	    If FontObj = Font Then
	      ' FontExists = True
	      FontExists = True
	    End If
	  Next
	  FontExists = False
	End Function
	Public Function GetStream() 'As String
	  Dim sItem
	  For Each sItem In m_streams
	    GetStream = sItem
	    m_streams.Remove sItem
	    Exit Function
	  Next
	End Function
	Public Function count() 'As Integer
	  count = m_streams.count
	End Function
	Public Property Get Resources() 'As String
	    Resources = m_resources
	End Property
	Public Property Let Resources(ByVal Value)
	    m_resources = Value
	End Property
End Class
'===================
Class Row
	Public default Property Get ClassName() 'As FontStyles
		ClassName = "Row"
	End Property
	Private m_cells 'As Scripting.Dictionary
	Private m_Height 'As Integer
	Private Sub Class_Initialize()
	 Set m_cells = CreateObject("Scripting.Dictionary")
	End Sub
	Public Sub AddCell(ByVal myCell)
	 Dim aCell 'As cell
	 Set aCell = myCell.GetCopy
	 m_cells.Add aCell, ""
	End Sub
	 Property Get HeightInPDFunits()
	 HeightInPDFunits = m_Height
	End Property
	 Property Get cells() 'As Scripting.Dictionary
	 Set cells = m_cells
	End Property
	function CalculateHeight(ByVal width, ByVal cellpadding)
	 Dim cell 'As cell
	 Dim H 'As Integer
	 Dim w 'As Integer' Printable width
	 m_Height = 0
	 width = width / m_cells.count
	 w = width - 2 * cellpadding
	 For Each cell In m_cells
	 H = cell.CalculateHeight(w)
	 if H > m_Height Then
	m_Height = H
	 End if
	 Next
	 m_Height = m_Height + 2 * cellpadding
	 CalculateHeight = m_Height
	End function
End Class
'===================
Class Table
	Public default Property Get ClassName() 'As FontStyles
		ClassName = "Table"
	End Property
	Private m_border 'As Borders
	Private m_rows 'As Scripting.Dictionary
	Private m_Height 'As Integer
	Public CellPaddingInPDFUnits 'As Integer
	Private m_ColumnWidth 'As Integer' PDF measurement
	Private m_cellCount 'As Integer
	Private m_ActualHeight 'As Integer
	Private m_startH 'As Integer
	Private m_StartV 'As Integer
	Private Sub Class_Initialize()
	 Set m_rows = CreateObject("Scripting.Dictionary")
	 m_border = Borders_thick
	 CellPaddingInPDFUnits = 4
	End Sub
	function GetCopy()
	End function
	Public Property Get StartPDFH() 'As Integer
	 StartPDFH = m_startH
	End Property
	Public Property Let StartPDFH(ByVal MyStartInPDFUnits)
	 m_startH = MyStartInPDFUnits
	End Property
	Public function Draw(ByRef StartV, ByVal width, ByRef FontAlias, _
	 ByRef pagenum, ByVal TopMargin) 'As PDFObject
	 Dim pdfObj 'As PDFObject
	 Dim row 'As row
	 Dim count 'As Integer
	 Dim TotalCols 'As Integer
	 Dim stream 'As String' Text Stream
	 Dim GStream 'As String' graphics stream
	 Dim cell 'As cell
	 Dim RightH 'As Integer
	 Dim V 'As Integer
	 Dim H 'As Integer
	 Dim RowStartV 'As Integer
	 Dim cols 'As Integer
	 Dim c 'As Integer
	 Dim accumColumn 'As Integer
	 Dim RowStarty 'As Integer
	 Set pdfObj = New PDFObject
	 Call CalculateTable(width)
	 ' Save start point
	 m_StartV = StartV
	 if m_border <> Borders_none Then
	 stream = "0.0 G " + vbCr ' Black color
	 if m_border = Borders_thick Then
	stream = "2 w " + vbCr ' Line width
	 Else
	stream = "1 w " + vbCr ' Line width
	 End if
	 RightH = m_startH + width
	 ' Top level line of the table
	 stream = stream + line(m_startH, StartV, RightH, StartV)
	 For Each row In m_rows
	' Print first vertical bar For Each cell
	V = StartV - row.HeightInPDFunits
	stream = stream + line(m_startH, StartV, m_startH, V)
	' Print right vertical bar For Each cell
	cols = 0
	c = 1
	accumColumn = 0
	For Each cell In row.cells
	cols = cols + cell.ColumnSpan
	if c = row.cells.count Then
	 H = RightH
	Else
	 if cell.WidthInPercent = 0 Then
	 H = m_startH + cols * m_ColumnWidth
	 Else
	 accumColumn = accumColumn + cell.WidthInPercent * width
	 H = m_startH + accumColumn
	 End if
	End if
	V = StartV - row.HeightInPDFunits
	if V < 1 Then Exit For
	stream = stream + line(H, StartV, H, V)
	c = c + 1
	Next
	' Print row divider
	StartV = StartV - row.HeightInPDFunits
	stream = stream + line(m_startH, StartV, RightH, StartV)
	 Next
	 End if
	 ' Print text in cells
	 V = m_StartV
	 For Each row In m_rows
	 H = m_startH
	 For Each cell In row.cells
	cell.StartPDFH = H + CellPaddingInPDFUnits
	cell.StartPDFV = V
	Set pdfObj = cell.Draw(FontAlias, 1, TopMargin)
	stream = stream + pdfObj.getStream
	if cell.WidthInPercent = 0 Then
	H = H + cell.ColumnSpan * m_ColumnWidth
	Else
	H = H + width * cell.WidthInPercent
	End if
	 Next
	 V = V - row.HeightInPDFunits
	 if V < 1 Then Exit For
	 Next
	 pdfObj.addStream (stream)
	 Set Draw = pdfObj
	End function
	Sub CalculateTable(ByVal width)
	 Dim row 'As row
	 m_ActualHeight = 0
	 ' Calculate table width
	 if width = 0 Then
	 Err.Raise 100,"","Zero Width table Not supported."
	 End if
	 ' Check To see that we have a column count
	 if m_rows.count < 1 Then
	 Err.Raise 100,"","No Rows To draw."
	 End if
	 m_cellCount = CalculateCellCount()
	 ' Column width when all columns have the same width
	 m_ColumnWidth = (width - 2 * m_border) / m_cellCount
	 ' Calculate
	 For Each row In m_rows
	 row.CalculateHeight m_ColumnWidth, CellPaddingInPDFUnits
	 m_ActualHeight = m_ActualHeight + row.HeightInPDFunits + 2 * m_border
	 Next
	 m_ActualHeight = m_ActualHeight + 2 * m_border
	End Sub
	Public Sub setColumnWidth(ByVal width)
	 ' This method sets the width of the table columns
	 ' Columns are from index 1 To the upper bound of width(). With(0) is Not used.
	 ' Each entry In the input array becomes a percentage of the sum of all entries in the input array
	 Dim row 'As row
	 Dim cell 'As cell
	 Dim totalWidth 'As Integer
	 Dim i 'As Integer
	 Dim cols 'As Integer
	 if m_rows.count < 1 Then
	 Err.Raise 100,"","No rows."
	 End if
	 m_cellCount = CalculateCellCount()
	 if m_cellCount <> UBound(width) Then
	 Err.Raise 100,"","Number of columns doesn't match the setting For column width."
	 End if
	 For i = 1 To UBound(width)
	 totalWidth = totalWidth + width(i) ' Calculate the total
	 Next
	 if totalWidth <= 0 Then
	 Err.Raise 100,"","Can't Set column width on table."
	 End if
	 For Each row In m_rows
	 cols = 0
	 For Each cell In row.cells
	cols = cols + cell.ColumnSpan
	cell.WidthInPercent = Math.Round(width(cols) / totalWidth, 2) ' Percent
	 Next
	 Next
	End Sub
	Private function CalculateCellCount() 'As Integer
	 Dim scellCnt 'As Integer
	 Dim cellCnt 'As Integer
	 Dim row 'As row
	 Dim cell 'As cell
	 For Each row In m_rows
	 cellCnt = 0
	 For Each cell In row.cells
	cellCnt = cellCnt + cell.ColumnSpan
	 Next
	 if scellCnt <> 0 And scellCnt <> cellCnt Then
	Err.Raise 100,"","Uneven number of cells With column span In the row collection."
	 End if
	 scellCnt = cellCnt
	 Next
	 if cellCnt = 0 Then
	 Err.Raise 100,"","No columns/cells."
	 End if
	 CalculateCellCount = cellCnt
	End function
	Private function line(ByVal x, ByVal y, ByVal x1, ByVal y1) 'As String
	 Dim stream 'As String
	 stream = stream & x & " " & y & " m" + vbCr
	 stream = stream & x1 & " " & y1 & " l" + vbCr
	 stream = stream & "S" + vbCr
	 line = stream
	End function
	Public Property Get Border() 'As Borders
	 Border = m_border
	End Property
	Public Property Let Border(ByVal myBorder)
	 Select Case myBorder
	 Case Borders_none
	m_border = myBorder
	 Case Borders_thick
	m_border = myBorder
	 Case Borders_thin
	m_border = myBorder
	 Case Else
	Err.Raise 100,"","Invalid Border"
	 End Select
	End Property
	Public Sub AddRow(ByVal myRow)
	 m_rows.Add myRow, ""
	End Sub
	Public function toString() 'As String
	 toString = "Table rows: " & m_rows.count
	End function
End Class
'===================
Class TextArea
	Public default Property Get ClassName() 'As FontStyles
		ClassName = "TextArea"
	End Property
	Private m_Texts 'As Scripting.Dictionary ' texts To be word wrapped
	Private m_LineQ 'As Scripting.Dictionary ' word wrapped lines
	Private m_StartV 'As Integer
	Private m_widthPDFUnits 'As Integer
	Public HeightInPDFunits 'As Integer
	Public StartPDFH 'As Integer
	Private Sub Class_Initialize()
	 Set m_Texts = CreateObject("Scripting.Dictionary")
	 StartPDFH = 72
	End Sub
	Sub CalculateHeight(ByVal width)
	 Dim myText 'As TextObject
	 Dim FontRef 'As String
	 Dim key 'As String
	 Dim sFontRef 'As String
	 Dim found 'As Boolean
	 Dim FontSize 'As Integer
	 Dim sFontSize 'As Integer
	 Dim i 'As Integer
	 Dim lineNo 'As Integer
	 Dim linelen 'As Integer
	 Dim textLine 'As TextObject
	 Dim line 'As String' Text line
	 Dim tmpline 'As String
	 Dim vspace 'As Integer
	 Dim ret 'As String
	 Dim fonto 'As FontObj
	 if width < 1 Then
	 Err.Raise 100,"","Invalid width For TextArea"
	 End if
	 m_widthPDFUnits = width
	 Set m_LineQ = CreateObject("Scripting.Dictionary")
	 ' Split the text up In lines
	 For Each myText In m_Texts
	 line = myText.Text
	 ' Escape PDF special characters ( and )
	 line = ReplaceText(ReplaceText(line, "(", "\("), ")", "\)")
	 line = Trim(line)
	 FontSize = myText.FontSize
	 linelen = myText.FontObj.HorizontalSpace * width / myText.FontSize
	 if Len(line) > linelen Then
	'word wrap
	Do While Len(line) > linelen
	tmpline = Left(line, linelen)
	For i = Len(tmpline) To Len(tmpline) / 2 Step -1
	 if InStr("*&^%$#,. ;<=>[])}!""", mid(tmpline, i, 1)) Then
	 ' find appropriate End of line
	 tmpline = Left(tmpline, i)
	 Exit For
	 End if
	Next
	line = mid(line, Len(tmpline) + 1)
	Set textLine = New TextObject
	With textLine
	 .Text = tmpline
	 Set .FontObj = myText.FontObj
	 .FontSize = myText.FontSize
	End With
	m_LineQ.Add textLine, ""
	Loop
	Set textLine = New TextObject
	With textLine
	.Text = line
	Set .FontObj = myText.FontObj
	.FontSize = myText.FontSize
	End With
	m_LineQ.Add textLine, ""
	 Else
	Set textLine = New TextObject
	With textLine
	.Text = line
	Set .FontObj = myText.FontObj
	.FontSize = myText.FontSize
	End With
	m_LineQ.Add textLine, ""
	 End if
	 Next
	 HeightInPDFunits = 0
	 For Each myText In m_LineQ
	 FontSize = myText.FontSize
	 HeightInPDFunits = HeightInPDFunits + 1.2 * FontSize
	 Next
	End Sub
	function Draw(ByRef StartV, ByVal width, ByRef FontAlias, _
	 ByRef pagenum, ByVal TopStart) 'As PDFObject
	 Dim PDFO 'As PDFObject
	 Dim myText 'As TextObject
	 Dim FontName 'As String
	 Dim TempPdfo 'As PDFObject
	 Dim FontRef 'As String
	 Dim key 'As String
	 Dim sFontRef 'As String
	 Dim found 'As Boolean
	 Dim FontSize 'As Integer
	 Dim sFontSize 'As Integer
	 Dim i 'As Integer
	 Dim lineNo 'As Integer
	 Dim linelen 'As Integer
	 Dim textLine 'As TextObject
	 Dim line 'As String' Text line
	 Dim tmpline 'As String
	 Dim vspace 'As Integer
	 Dim ret 'As String
	 Dim fonto 'As FontObj
	 Dim page 'As PageBreak
	 Dim save 'As String
	 Call CalculateHeight(width)
	 Set PDFO = New PDFObject
	 Set page = New PageBreak
	 ' Process fonts
	 For Each myText In m_Texts
	 ret = myText.Text
	 ' Set if we have this font
	 FontRef = getFontNumber(myText.Font, myText.FontStyle, FontAlias)
	 if FontRef = "" Then
	'Add a new font
	FontRef = Trim(CStr(FontAlias.count + 1))
	Set fonto = New CFontObj
	With fonto
	.FontRef = FontRef
	.Font = myText.Font
	.FontStyle = myText.FontStyle
	End With
	FontAlias.Add FontRef, fonto
	 End if
	 myText.FontObj.FontRef = FontRef
	 found = False
	 For Each key In PDFO.m_fonts
	if key = FontRef Then found = True
	 Next
	 if found = False Then
	if Not PDFO.m_fonts.Exists(FontRef) Then
	PDFO.m_fonts.Add FontRef, ""
	End if
	 End if
	 Next
	 ' Print the lines To the PDF document
	 lineNo = -1
	 ret = " BT" + vbCr ' Begin text object
	 For Each myText In m_LineQ
	 line = myText.Text
	 FontName = myText.FontObj.FontName()
	 FontRef = myText.FontObj.FontRef
	 FontSize = myText.FontSize
	 vspace = 1.2 * FontSize
	 if (sFontRef <> FontRef) Or sFontSize <> FontSize Then
	ret = ret + "/F" & FontRef & " " & FontSize & " Tf" & vbCr ' Text and font
	ret = ret + "1 0 0 1 " & StartPDFH & " " & StartV & " Tm" & vbCr ' Set text matrix
	ret = ret + CStr(vspace) & " TL" & vbCr ' Set text leading
	'lineNo = lineNo + 1
	 End if
	 sFontRef = FontRef
	 sFontSize = FontSize
	 ret = ret + "T* (" & line & vbCrLf & ") Tj" & vbCr
	 StartV = StartV - vspace
	 if StartV < 100 Then
	' Print footer
	page.StartPDFH = StartPDFH
	ret = ret + "ET " + vbCrLf
	Set TempPdfo = page.Draw(StartV, width, FontAlias, pagenum, TopStart)
	ret = ret + TempPdfo.getStream()
	PDFO.addStream (ret)
	PDFO.PageBreak = True
	' Start new page
	save = ret
	ret = ""
	ret = "BT " + vbCrLf
	sFontRef = ""
	StartV = TopStart
	 End if
	 Next
	 StartV = StartV - vspace
	 ret = ret + " ET" + vbCr
	 PDFO.addStream (ret)
	 Set Draw = PDFO
	End function
	function GetCopy()
	 Dim Text 'As TextArea
	 Dim tobj 'As TextObject
	 Set Text = New TextArea
	 For Each tobj In m_Texts
	 Text.AddText tobj.Text, tobj.Font, tobj.FontSize, tobj.FontStyle
	 Next
	 With Text
	 .StartPDFH = StartPDFH
	 End With
	 Set GetCopy = Text
	End function
	function getTexts() 'As Scripting.Dictionary
	 Set getTexts = m_Texts
	End function
	Public Sub AddText(ByVal Text, ByVal Font, ByVal FontSize, ByVal style)
	 if Font = "" Then Font = Fonts_Helvetica
	 if FontSize = "" Then FontSize = 10
	 if CStr(style) = "" Then style = FontStyles_Regular
	 Dim myText 'As TextObject
	 Set myText = New TextObject
	 With myText
	 .Font = Font
	 .FontSize = FontSize
	 .Text = Text
	 .FontStyle = style
	 End With
	 m_Texts.Add myText, ""
	End Sub
	 function toString() 'As String
	 Dim ret 'As String
	 ret = "TextArea: "
	 if m_Texts.count > 0 Then
	 ret = ret + GetDictionaryItem(m_Texts, 1).Text
	 End if
	 toString = ret
	End function
	function GetDictionaryItem(dic, ByVal iIndex)
	 Dim oItem, i
	 i = 0
	 For Each oItem In dic
	 i = i + 1
	 if i = iIndex Then
	if IsObject(oItem) Then
	Set GetDictionaryItem = oItem
	Else
	GetDictionaryItem = oItem
	End if
	Exit function
	 End if
	 Next
	End function
	Private function getFontNumber(ByVal Font, _
	ByVal FontStyle, _
	ByRef Fonts) 'As String
	 Dim i 'As Integer
	 Dim key 'As String
	 Dim fName 'As String
	 Dim fonto 'As FontObj
	 For i = 1 To Fonts.count
	 key = Trim(CStr(i))
	 Set fonto = Fonts(key)
	 if fonto.Font = Font And fonto.FontStyle = FontStyle Then
	'If Font.equals(fonto) Then
	getFontNumber = fonto.FontRef
	 End if
	 Next
	End function
	Public function ReplaceText(ByRef Text_Renamed, ByRef TextToReplace, ByRef NewText) 'As String
	 Dim mtext 'As String
	 Dim SpacePos 'As Integer
	 mtext = Text_Renamed
	 SpacePos = InStr(mtext, TextToReplace)
	 Do While SpacePos
	 mtext = Left(mtext, SpacePos) & NewText & mid(mtext, SpacePos + Len(TextToReplace))
	 SpacePos = InStr(SpacePos + Len(NewText), mtext, TextToReplace)
	 Loop
	 ReplaceText = mtext
	End function
End Class
'===================
Class TextObject
	Public default Property Get ClassName() 'As FontStyles
		ClassName = "TextObject"
	End Property
	Dim m_Text 'As String
	Dim m_Font 'As FontObj
	Public FontSize 'As Integer
	Private Sub Class_Initialize()
	 Set m_Font = New CFontObj
	 FontSize = 10
	End Sub
	Public Property Get FontStyle() 'As FontStyles
	 FontStyle = FontObj.FontStyle
	End Property
	Public Property Let FontStyle(ByVal MyStyle)
	 m_Font.FontStyle = MyStyle
	End Property
	Public Property Get Font() 'As Fonts
	 Font = m_Font.Font
	End Property
	Public Property Let Font(ByVal myFont)
	 m_Font.Font = myFont
	End Property
	Public Property Get Text() 'As String
	 Text = m_Text
	End Property
	Public Property Let Text(ByVal myText)
	 m_Text = myText
	End Property
	Public Property Get FontObj() 'As FontObj
	 Set FontObj = m_Font
	End Property
	Public Property Set FontObj(ByVal myFont)
	 Set m_Font = myFont
	End Property
End Class
'===================
Function GetPdfFormatedDate()
	GetPdfFormatedDate = year(Now) & _
		PadLeftWithZeros(month(now),2) & _
		PadLeftWithZeros(day(now),2) & _
		PadLeftWithZeros(hour(now),2) & _
		PadLeftWithZeros(minute(now),2) & _
		PadLeftWithZeros(second(now),2)
End function
Function PadLeftWithZeros(sValue,iSize)
	PadLeftWithZeros = right("00000000" + trim(sValue),iSize)
	End function
	function StringToMultiByte(S)
	Dim i, MultiByte
	For i=1 To Len(S)
	MultiByte = MultiByte & ChrB(Asc(Mid(S,i,1)))
	Next
	StringToMultiByte = MultiByte
End function
%>
```

