Imports MigraDoc.DocumentObjectModel
Imports MigraDoc.DocumentObjectModel.Shapes
Imports MigraDoc.DocumentObjectModel.Tables
'Imports MigraDoc.Rendering

'Namespace Certificate
Public Class Certificate
    'The document to be rendered as pdf
    Public document As Document

    Dim TableBorder_Black As Color = New Color(0, 0, 0)
    Dim CellBackgroung_GBlue As Color = New Color(218, 218, 225)

    Dim PicL As Image
    'Text Frames to add address and other info to the page
    Dim AboveTable1 As TextFrame
    Dim AboveTable3 As TextFrame
    Dim BelowTable3 As TextFrame
    Dim AboveTable4 As TextFrame
    Dim AboveTable5 As TextFrame
    Dim BelowTable5 As TextFrame

    'Dim myLine As LineFormat = New LineFormat()

    'To hold the table of elements in the certification
    Public table1FirstPage As Table
    Public table2FirstPage As Table
    Public table3FirstPage As Table
    Public table4SecondPage As Table
    Public table5ThirdPage As Table

    Private strOrderNumber As String
    'Private myDataGridView As System.Windows.Forms.DataGridView
    Private myFM As String
    Private myRevision As String

    Public Sub New(ByVal pStrOrderNum As String, ByVal pRevision As String, ByVal pFM As String)
        strOrderNumber = pStrOrderNum
        'myDataGridView = pGrid
        myRevision = pRevision
        myFM = pFM
    End Sub

    Public Function CreateDocument() As Document


        'Create a new MigraDoc document
        Me.document = New Document()
        Me.document.Info.Title = "Raw Material Incoming Inspection Report"
        Me.document.Info.Author = "CareFusion"
        Me.document.Info.Subject = "Reporte de inspección de material"
        Me.document.Info.Keywords = "CareFusion, Report, Inspección, material"


        'Call the sub to define styles
        DefineStyles()

        'Call Create Page
        CreatePage()

        Return Me.document
    End Function

    '' Defines the styles used to format the MigraDoc document.

    Public Sub DefineStyles()
        ''Get the predefined style Normal.
        Dim style As Style = Me.document.Styles("Normal")
        ''Because all styles are derived from Normal, the next line changes the 
        '' font of the whole document. Or, more exactly, it changes the font of
        '' all styles and paragraphs that do not redefine the font.
        style.Font.Name = "Arial"

        style = Me.document.Styles(StyleNames.Header)
        style.ParagraphFormat.AddTabStop("16cm", TabAlignment.Right)

        style = Me.document.Styles(StyleNames.Footer)
        style.ParagraphFormat.AddTabStop("8cm", TabAlignment.Center)

        '' Create a new style called Table based on style Normal
        style = Me.document.Styles.AddStyle("Table", "Normal")
        style.Font.Name = "Times New Roman"
        style.Font.Name = "Arial"
        style.Font.Size = 9

        '' Create a new style called Reference based on style Normal
        style = Me.document.Styles.AddStyle("Reference", "Normal")
        style.ParagraphFormat.SpaceBefore = "5mm"
        style.ParagraphFormat.SpaceAfter = "5mm"
        style.ParagraphFormat.TabStops.AddTabStop("16cm", TabAlignment.Right)
    End Sub

    Public Sub CreatePage()

        '' Add a section to the document
        Dim section As Section = Me.document.AddSection()

        'Top and bottom margins
        section.PageSetup.BottomMargin = "0.889cm"
        section.PageSetup.TopMargin = "0.254cm"
        section.PageSetup.RightMargin = "1cm"
        section.PageSetup.LeftMargin = "1cm"

        'Create the footer
        Dim paragraph As Paragraph = section.Footers.Primary.AddParagraph
        paragraph.Format.Font.Name = "Times New Roman"
        paragraph.AddText("Form-085")
        paragraph.AddSpace(43)
        paragraph.AddText("Rev. J (01/29/2013)")
        paragraph.AddTab()
        paragraph.AddTab()
        paragraph.AddTab()
        paragraph.AddText("CAF# 2012-09-08")
        paragraph.AddTab()
        paragraph.AddText("Page ")
        paragraph.AddPageField()
        paragraph.AddText(" of ")
        paragraph.AddNumPagesField()
        paragraph.Format.Font.Size = 10
        paragraph.Format.Alignment = ParagraphAlignment.Center

        'Create the header
        PicL = section.Headers.Primary.AddImage("C:\Users\Kevin\Documents\Visual Studio 2015\Projects\WindowsForm PdfSharp Recreated\WindowsForm PdfSharp Recreated\Logo_Bard.png")
        PicL.Height = "1.7cm"
        PicL.LockAspectRatio = True
        PicL.RelativeVertical = RelativeVertical.Margin
        PicL.RelativeHorizontal = RelativeHorizontal.Margin
        PicL.Top = ShapePosition.Top
        PicL.Left = ShapePosition.Left
        PicL.WrapFormat.Style = WrapStyle.Through

        paragraph = section.Headers.Primary.AddParagraph
        paragraph.AddFormattedText("Reference WI-2.11.2A", TextFormat.Italic)
        paragraph.AddLineBreak()
        paragraph.Format.Font.Size = 9
        paragraph.Format.Alignment = ParagraphAlignment.Right

        '**********************************
        '*-----------FIRST PAGE-----------*
        '**********************************

        'Add the address text frame
        Me.AboveTable1 = section.AddTextFrame()
        Me.AboveTable1.Width = "13cm"
        Me.AboveTable1.Left = ShapePosition.Center
        Me.AboveTable1.RelativeHorizontal = RelativeHorizontal.Margin
        Me.AboveTable1.RelativeVertical = RelativeVertical.Page

        ' Writing what's above the first table
        paragraph = Me.AboveTable1.AddParagraph()
        paragraph.AddLineBreak()
        paragraph.AddLineBreak()
        paragraph.AddLineBreak()
        paragraph.AddTab()
        paragraph.AddText("Raw Material Incoming Inspection Report")
        paragraph.Format.Font.Size = 16
        paragraph.AddLineBreak()
        For i As Integer = 0 To 6
            paragraph = section.AddParagraph()
        Next

        'create Table
        Me.table1FirstPage = section.AddTable()
        Me.table1FirstPage.Style = "Table"
        Me.table1FirstPage.Borders.Color = TableBorder_Black
        Me.table1FirstPage.Borders.Width = 0.25
        Me.table1FirstPage.Borders.Left.Width = 0.7
        Me.table1FirstPage.Borders.Right.Width = 0.7
        Me.table1FirstPage.Rows.LeftIndent = 15.1


        'Create the columns of the first table
        '17.941cm

        Dim column As Column = Me.table1FirstPage.AddColumn("3.535cm")
        column.Format.Alignment = ParagraphAlignment.Center

        column = Me.table1FirstPage.AddColumn("5.436cm")
        column.Format.Alignment = ParagraphAlignment.Center

        column = Me.table1FirstPage.AddColumn("3.535cm")
        column.Format.Alignment = ParagraphAlignment.Center

        column = Me.table1FirstPage.AddColumn("5.436cm")
        column.Format.Alignment = ParagraphAlignment.Center

        'Create the header of the first table
        Dim row As Row = table1FirstPage.AddRow()
        row.HeadingFormat = True
        row.Format.Alignment = ParagraphAlignment.Center
        row.Format.Font.Size = 11
        row.Height = 21
        row.Cells(0).AddParagraph("RMI                         Ship to Stock ")
        row.Cells(0).VerticalAlignment = VerticalAlignment.Center
        row.Cells(0).Shading.Color = CellBackgroung_GBlue
        row.Cells(0).MergeRight = 3

        'For i As Integer = 1 To 5
        '    row = table.AddRow()
        '    row.Format.Alignment = ParagraphAlignment.Center
        '    row.Cells(0).AddParagraph("Col " & i)
        'Next

        row = table1FirstPage.AddRow()
        '------------ FIRST ROW OF DATA -----------
        row.Height = 20.67
        row.Cells(0).Format.Font.Size = 11
        row.Cells(0).AddParagraph("Supplier")
        row.Cells(0).Format.Alignment = ParagraphAlignment.Right
        row.Cells(0).VerticalAlignment = VerticalAlignment.Center
        row.Cells(0).Shading.Color = CellBackgroung_GBlue
        '-----
        row.Cells(1).AddParagraph(" ")
        '-----
        row.Cells(2).Format.Font.Size = 11
        row.Cells(2).AddParagraph("Data Recieved")
        row.Cells(2).Format.Alignment = ParagraphAlignment.Right
        row.Cells(2).VerticalAlignment = VerticalAlignment.Center
        row.Cells(2).Shading.Color = CellBackgroung_GBlue
        '-----
        row.Cells(3).AddParagraph(" ")
        '--------------------------------------------------------

        row = table1FirstPage.AddRow()
        '----------- SECOND ROW OF DATA ---------
        row.Height = 20.67
        row.Cells(0).Format.Font.Size = 11
        row.Cells(0).AddParagraph("PO #")
        row.Cells(0).Format.Alignment = ParagraphAlignment.Right
        row.Cells(0).VerticalAlignment = VerticalAlignment.Center
        row.Cells(0).Shading.Color = CellBackgroung_GBlue
        '-----
        row.Cells(1).AddParagraph(" ")
        '-----
        row.Cells(2).Format.Font.Size = 11
        row.Cells(2).AddParagraph("Reciever No")
        row.Cells(2).Format.Alignment = ParagraphAlignment.Right
        row.Cells(2).VerticalAlignment = VerticalAlignment.Center
        row.Cells(2).Shading.Color = CellBackgroung_GBlue
        '-----
        row.Cells(3).AddParagraph(" ")

        '--------------------------------------------------------

        row = table1FirstPage.AddRow()
        '----------- THIRD ROW OF DATA ---------
        row.Height = 25.67
        row.Cells(0).Format.Font.Size = 11
        row.Cells(0).AddParagraph("Manufacturing /")
        row.Cells(0).AddParagraph("Expiration Date")
        row.Cells(0).Format.Alignment = ParagraphAlignment.Right
        row.Cells(0).VerticalAlignment = VerticalAlignment.Center
        row.Cells(0).Shading.Color = CellBackgroung_GBlue
        '-----
        row.Cells(1).AddParagraph(" ")
        '-----
        row.Cells(2).Format.Font.Size = 11
        row.Cells(2).AddParagraph("CareFusion /")
        row.Cells(2).AddParagraph(" Vendor Lot No /")
        row.Cells(2).Format.Alignment = ParagraphAlignment.Right
        row.Cells(2).VerticalAlignment = VerticalAlignment.Center
        row.Cells(2).Shading.Color = CellBackgroung_GBlue
        '-----
        row.Cells(3).AddParagraph(" ")

        '--------------------------------------------------------

        row = table1FirstPage.AddRow()
        '----------- FOURTH ROW OF DATA ---------
        row.Height = 20
        row.Cells(0).Format.Font.Size = 11
        row.Cells(0).AddParagraph("Part No")
        row.Cells(0).Format.Alignment = ParagraphAlignment.Right
        row.Cells(0).VerticalAlignment = VerticalAlignment.Center
        row.Cells(0).Shading.Color = CellBackgroung_GBlue
        '-----
        row.Cells(1).AddParagraph(" ")
        '-----
        row.Cells(2).Format.Font.Size = 11
        row.Cells(2).AddParagraph("RMS Revision")
        row.Cells(2).Format.Alignment = ParagraphAlignment.Right
        row.Cells(2).VerticalAlignment = VerticalAlignment.Center
        row.Cells(2).Shading.Color = CellBackgroung_GBlue
        '-----
        row.Cells(3).AddParagraph(" ")

        '--------------------------------------------------------

        row = table1FirstPage.AddRow()
        '----------- FIFTH ROW OF DATA ---------
        row.Height = 26
        row.Cells(0).Format.Font.Size = 11
        row.Cells(0).AddParagraph("Quantity")
        row.Cells(0).Format.Alignment = ParagraphAlignment.Right
        row.Cells(0).VerticalAlignment = VerticalAlignment.Center
        row.Cells(0).Shading.Color = CellBackgroung_GBlue
        '-----
        row.Cells(1).AddParagraph(" ")
        '-----
        row.Cells(2).Format.Font.Size = 11
        row.Cells(2).AddParagraph("Containers Recieved")
        row.Cells(2).Format.Alignment = ParagraphAlignment.Right
        row.Cells(2).VerticalAlignment = VerticalAlignment.Center
        row.Cells(2).Shading.Color = CellBackgroung_GBlue
        '-----
        row.Cells(3).AddParagraph(" ")

        '--------------------------------------------------------

        row = table1FirstPage.AddRow()
        '----------- SIXTH ROW OF DATA ---------
        row.Height = 25.27
        row.Cells(0).Format.Font.Size = 11
        row.Cells(0).AddParagraph("Unit of Measure")
        row.Cells(0).Format.Alignment = ParagraphAlignment.Right
        row.Cells(0).VerticalAlignment = VerticalAlignment.Center
        row.Cells(0).Shading.Color = CellBackgroung_GBlue
        '-----
        row.Cells(1).AddParagraph(" ")
        '-----
        row.Cells(2).Format.Font.Size = 11
        row.Cells(2).AddParagraph("Boxes or Roll #")
        row.Cells(2).AddParagraph("inspected")
        row.Cells(2).Format.Alignment = ParagraphAlignment.Right
        row.Cells(2).VerticalAlignment = VerticalAlignment.Center
        row.Cells(2).Shading.Color = CellBackgroung_GBlue
        '-----
        row.Cells(3).AddParagraph(" ")

        '--------------------------------------------------------
        row = table1FirstPage.AddRow()
        '----------- SEVENTH ROW OF DATA ---------
        row.Height = 20.27
        row.Cells(0).Format.Font.Size = 11
        row.Cells(0).AddParagraph("Resin No")
        row.Cells(0).Format.Alignment = ParagraphAlignment.Right
        row.Cells(0).VerticalAlignment = VerticalAlignment.Center
        row.Cells(0).Shading.Color = CellBackgroung_GBlue
        '-----
        row.Cells(1).AddParagraph(" ")
        '-----
        row.Cells(2).Format.Font.Size = 11
        row.Cells(2).AddParagraph("Inspection Level")
        row.Cells(2).Format.Alignment = ParagraphAlignment.Right
        row.Cells(2).VerticalAlignment = VerticalAlignment.Center
        row.Cells(2).Shading.Color = CellBackgroung_GBlue
        '-----
        row.Cells(3).Format.Font.Size = 11
        row.Cells(3).AddParagraph("Reduced  Normal  Tightened ")
        row.Cells(3).Format.Alignment = ParagraphAlignment.Center
        row.Cells(3).VerticalAlignment = VerticalAlignment.Center

        '--------------------------------------------------------

        row = table1FirstPage.AddRow()
        '----------- EIGTH ROW OF DATA ---------
        row.Height = 20.67
        row.Cells(0).Format.Font.Size = 11
        row.Cells(0).AddParagraph("Sampled By")
        row.Cells(0).Format.Alignment = ParagraphAlignment.Right
        row.Cells(0).VerticalAlignment = VerticalAlignment.Center
        row.Cells(0).Shading.Color = CellBackgroung_GBlue
        '-----
        row.Cells(1).AddParagraph(" ")
        '-----
        row.Cells(2).Format.Font.Size = 11
        row.Cells(2).AddParagraph("Data Sampled")
        row.Cells(2).Format.Alignment = ParagraphAlignment.Right
        row.Cells(2).VerticalAlignment = VerticalAlignment.Center
        row.Cells(2).Shading.Color = CellBackgroung_GBlue
        '-----
        row.Cells(3).AddParagraph(" ")

        '-------END OF FIRST TABLE------
        paragraph = section.AddParagraph()
        '--------------------------------------------------------

        '-------START OF SECOND TABLE------

        'create the second Table
        Me.table2FirstPage = section.AddTable()
        Me.table2FirstPage.Style = "Table"
        Me.table2FirstPage.Borders.Color = TableBorder_Black
        Me.table2FirstPage.Borders.Width = 0.25
        Me.table2FirstPage.Borders.Left.Width = 0.7
        Me.table2FirstPage.Borders.Right.Width = 0.7
        Me.table2FirstPage.Rows.LeftIndent = 15.1

        'Create the columns of the second table
        '17.941cm

        Dim Table2Column As Column = Me.table2FirstPage.AddColumn("2.243cm")
        Table2Column.Format.Alignment = ParagraphAlignment.Center

        Table2Column = Me.table2FirstPage.AddColumn("2.243cm")
        Table2Column.Format.Alignment = ParagraphAlignment.Center

        Table2Column = Me.table2FirstPage.AddColumn("1.543cm")
        Table2Column.Format.Alignment = ParagraphAlignment.Center

        Table2Column = Me.table2FirstPage.AddColumn("1.543cm")
        Table2Column.Format.Alignment = ParagraphAlignment.Center

        Table2Column = Me.table2FirstPage.AddColumn("2.746cm")
        Table2Column.Format.Alignment = ParagraphAlignment.Center

        Table2Column = Me.table2FirstPage.AddColumn("2.243cm")
        Table2Column.Format.Alignment = ParagraphAlignment.Center

        Table2Column = Me.table2FirstPage.AddColumn("3.143cm")
        Table2Column.Format.Alignment = ParagraphAlignment.Center

        Table2Column = Me.table2FirstPage.AddColumn("2.243cm")
        Table2Column.Format.Alignment = ParagraphAlignment.Center

        'Create the header of the second table
        Dim Table2Row As Row = table2FirstPage.AddRow()
        Table2Row.HeadingFormat = False
        Table2Row.Format.Font.Size = 11
        Table2Row.Format.Alignment = ParagraphAlignment.Center
        '-----
        Table2Row.Cells(0).AddParagraph("Defect Class")
        Table2Row.Cells(0).MergeDown = 2
        Table2Row.Cells(0).Format.Alignment = ParagraphAlignment.Center
        Table2Row.Cells(0).VerticalAlignment = VerticalAlignment.Center
        Table2Row.Cells(0).Shading.Color = CellBackgroung_GBlue
        '-----
        Table2Row.Cells(1).AddParagraph("ANSI/ASQ z1.4 Sampling Level")
        Table2Row.Cells(1).MergeDown = 2
        Table2Row.Cells(1).Format.Alignment = ParagraphAlignment.Center
        Table2Row.Cells(1).VerticalAlignment = VerticalAlignment.Center
        Table2Row.Cells(1).Shading.Color = CellBackgroung_GBlue
        '-----
        Table2Row.Cells(2).AddParagraph("AQL")
        Table2Row.Cells(2).MergeDown = 2
        Table2Row.Cells(2).Format.Alignment = ParagraphAlignment.Center
        Table2Row.Cells(2).VerticalAlignment = VerticalAlignment.Center
        Table2Row.Cells(2).Shading.Color = CellBackgroung_GBlue
        '-----
        Table2Row.Cells(3).AddParagraph("Sample Size")
        Table2Row.Cells(3).MergeDown = 2
        Table2Row.Cells(3).Format.Alignment = ParagraphAlignment.Center
        Table2Row.Cells(3).VerticalAlignment = VerticalAlignment.Center
        Table2Row.Cells(3).Shading.Color = CellBackgroung_GBlue
        '-----
        Table2Row.Cells(4).AddParagraph("Accept/")
        Table2Row.Cells(4).AddParagraph("Reject")
        Table2Row.Cells(4).MergeDown = 2
        Table2Row.Cells(4).Format.Alignment = ParagraphAlignment.Center
        Table2Row.Cells(4).VerticalAlignment = VerticalAlignment.Center
        Table2Row.Cells(4).Shading.Color = CellBackgroung_GBlue
        '-----
        Table2Row.Cells(5).AddParagraph("No. of Defects Found")
        Table2Row.Cells(5).MergeDown = 2
        Table2Row.Cells(5).Format.Alignment = ParagraphAlignment.Center
        Table2Row.Cells(5).VerticalAlignment = VerticalAlignment.Center
        Table2Row.Cells(5).Shading.Color = CellBackgroung_GBlue
        '-----
        Table2Row.Cells(6).AddParagraph("Type of Defect")
        Table2Row.Cells(6).AddParagraph("(if any)")
        Table2Row.Cells(6).MergeDown = 2
        Table2Row.Cells(6).Format.Alignment = ParagraphAlignment.Center
        Table2Row.Cells(6).VerticalAlignment = VerticalAlignment.Center
        Table2Row.Cells(6).Shading.Color = CellBackgroung_GBlue
        '-----
        Table2Row.Cells(7).AddParagraph("Pass /")
        Table2Row.Cells(7).AddParagraph("Fail")
        Table2Row.Cells(7).MergeDown = 2
        Table2Row.Cells(7).Format.Alignment = ParagraphAlignment.Center
        Table2Row.Cells(7).VerticalAlignment = VerticalAlignment.Center
        Table2Row.Cells(7).Shading.Color = CellBackgroung_GBlue

        Table2Row = table2FirstPage.AddRow()
        Table2Row = table2FirstPage.AddRow()
        '------ END OF HEADER ------

        '------ First Row of Data -----
        Table2Row = table2FirstPage.AddRow()
        Table2Row.Height = 32
        Table2Row.Cells(0).AddParagraph("Class II")
        Table2Row.Cells(0).AddParagraph("(Variable)")
        Table2Row.Cells(0).MergeDown = 2
        Table2Row.Cells(0).Format.Font.Size = 11
        Table2Row.Cells(0).Format.Alignment = ParagraphAlignment.Center
        Table2Row.Cells(0).VerticalAlignment = VerticalAlignment.Center
        Table2Row.Cells(0).Shading.Color = CellBackgroung_GBlue

        For i As Integer = 1 To 7
            If i = 4 Then
                Table2Row.Cells(i).Format.Font.Size = 11
                Table2Row.Cells(i).AddParagraph("a =      r =    ")
                Table2Row.Cells(i).MergeDown = 2
            Else
                Table2Row.Cells(i).AddParagraph()
                Table2Row.Cells(i).MergeDown = 2
            End If
            Table2Row.Format.Alignment = ParagraphAlignment.Center
            Table2Row.Cells(i).VerticalAlignment = VerticalAlignment.Center
        Next

        Table2Row = table2FirstPage.AddRow()
        Table2Row = table2FirstPage.AddRow()

        '------ Second Row of Data -----
        Table2Row = table2FirstPage.AddRow()
        Table2Row.Height = 32
        Table2Row.Cells(0).AddParagraph("Class II")
        Table2Row.Cells(0).AddParagraph("(Visual)")
        Table2Row.Cells(0).MergeDown = 2
        Table2Row.Cells(0).Format.Font.Size = 11
        Table2Row.Cells(0).Format.Alignment = ParagraphAlignment.Center
        Table2Row.Cells(0).VerticalAlignment = VerticalAlignment.Center
        Table2Row.Cells(0).Shading.Color = CellBackgroung_GBlue

        For i As Integer = 1 To 7
            If i = 4 Then
                Table2Row.Cells(i).Format.Font.Size = 11
                Table2Row.Cells(i).AddParagraph("a =      r =    ")
                Table2Row.Cells(i).MergeDown = 2
            Else
                Table2Row.Cells(i).AddParagraph()
                Table2Row.Cells(i).MergeDown = 2
            End If
            Table2Row.Format.Alignment = ParagraphAlignment.Center
            Table2Row.Cells(i).VerticalAlignment = VerticalAlignment.Center
        Next

        Table2Row = table2FirstPage.AddRow()
        Table2Row = table2FirstPage.AddRow()

        '------ Third Row of Data -----
        Table2Row = table2FirstPage.AddRow()
        Table2Row.Height = 32
        Table2Row.Cells(0).AddParagraph("Class III")
        Table2Row.Cells(0).AddParagraph("(Variable)")
        Table2Row.Cells(0).MergeDown = 2
        Table2Row.Cells(0).Format.Font.Size = 11
        Table2Row.Cells(0).Format.Alignment = ParagraphAlignment.Center
        Table2Row.Cells(0).VerticalAlignment = VerticalAlignment.Center
        Table2Row.Cells(0).Shading.Color = CellBackgroung_GBlue

        For i As Integer = 1 To 7
            If i = 4 Then
                Table2Row.Cells(i).Format.Font.Size = 11
                Table2Row.Cells(i).AddParagraph("a =      r =    ")
                Table2Row.Cells(i).MergeDown = 2
            Else
                Table2Row.Cells(i).AddParagraph()
                Table2Row.Cells(i).MergeDown = 2
            End If
            Table2Row.Format.Alignment = ParagraphAlignment.Center
            Table2Row.Cells(i).VerticalAlignment = VerticalAlignment.Center
        Next

        Table2Row = table2FirstPage.AddRow()
        Table2Row = table2FirstPage.AddRow()

        '------ Fourth Row of Data -----
        Table2Row = table2FirstPage.AddRow()
        Table2Row.Height = 32
        Table2Row.Cells(0).AddParagraph("Class III")
        Table2Row.Cells(0).AddParagraph("(Visual)")
        Table2Row.Cells(0).MergeDown = 2
        Table2Row.Cells(0).Format.Font.Size = 11
        Table2Row.Cells(0).Format.Alignment = ParagraphAlignment.Center
        Table2Row.Cells(0).VerticalAlignment = VerticalAlignment.Center
        Table2Row.Cells(0).Shading.Color = CellBackgroung_GBlue

        For i As Integer = 1 To 7
            If i = 4 Then
                Table2Row.Cells(i).Format.Font.Size = 11
                Table2Row.Cells(i).AddParagraph("a =      r =    ")
                Table2Row.Cells(i).MergeDown = 2
            Else
                Table2Row.Cells(i).AddParagraph()
                Table2Row.Cells(i).MergeDown = 2
            End If
            Table2Row.Format.Alignment = ParagraphAlignment.Center
            Table2Row.Cells(i).VerticalAlignment = VerticalAlignment.Center
        Next

        Table2Row = table2FirstPage.AddRow()
        Table2Row = table2FirstPage.AddRow()

        '------ Fifth Row of Data -----
        Table2Row = table2FirstPage.AddRow()
        Table2Row.Height = 32
        Table2Row.Cells(0).AddParagraph("Class IV")
        Table2Row.Cells(0).AddParagraph("(Visual)")
        Table2Row.Cells(0).MergeDown = 2
        Table2Row.Cells(0).Format.Font.Size = 11
        Table2Row.Cells(0).Format.Alignment = ParagraphAlignment.Center
        Table2Row.Cells(0).VerticalAlignment = VerticalAlignment.Center
        Table2Row.Cells(0).Shading.Color = CellBackgroung_GBlue

        For i As Integer = 1 To 7
            If i = 4 Then
                Table2Row.Cells(i).Format.Font.Size = 11
                Table2Row.Cells(i).AddParagraph("a =      r =    ")
                Table2Row.Cells(i).MergeDown = 2
            Else
                Table2Row.Cells(i).AddParagraph()
                Table2Row.Cells(i).MergeDown = 2
            End If
            Table2Row.Format.Alignment = ParagraphAlignment.Center
            Table2Row.Cells(i).VerticalAlignment = VerticalAlignment.Center
        Next

        Table2Row = table2FirstPage.AddRow()
        Table2Row = table2FirstPage.AddRow()
        '--------------END OF SECOND TABLE-------------

        paragraph = section.AddParagraph()
        '------------START OF THIRD TABLE--------------

        'Add the text frame
        Me.AboveTable3 = section.AddTextFrame()
        Me.AboveTable3.Width = "18cm"
        Me.AboveTable3.Left = ShapePosition.Center
        Me.AboveTable3.RelativeHorizontal = RelativeHorizontal.Margin
        Me.AboveTable3.RelativeVertical = RelativeVertical.Page

        ' Writing what's above the third table
        paragraph = Me.AboveTable3.AddParagraph()
        For c As Integer = 1 To 42
            paragraph.AddLineBreak()
        Next
        paragraph.Format.Font.Size = 11
        paragraph.AddSpace(3)
        paragraph.AddFormattedText("c", New Font("Webdings"))
        paragraph.AddFormattedText(" Lot # inspected from previous shipment sample size increased by __________ samples.  ")
        paragraph.AddLineBreak()
        paragraph = section.AddParagraph()
        paragraph = section.AddParagraph()

        'create the third Table
        Me.table3FirstPage = section.AddTable()
        Me.table3FirstPage.Style = "Table"
        Me.table3FirstPage.Borders.Color = TableBorder_Black
        Me.table3FirstPage.Borders.Width = 0.25
        Me.table3FirstPage.Borders.Left.Width = 0.7
        Me.table3FirstPage.Borders.Right.Width = 0.7
        Me.table3FirstPage.Rows.LeftIndent = 15.1

        'Create the columns of the third table
        '17.941cm

        Dim Table3Column As Column = Me.table3FirstPage.AddColumn("4.488cm")
        Table3Column.Format.Alignment = ParagraphAlignment.Center

        Table3Column = Me.table3FirstPage.AddColumn("4.488cm")
        Table3Column.Format.Alignment = ParagraphAlignment.Center

        Table3Column = Me.table3FirstPage.AddColumn("4.488cm")
        Table3Column.Format.Alignment = ParagraphAlignment.Center

        Table3Column = Me.table3FirstPage.AddColumn("4.488cm")
        Table3Column.Format.Alignment = ParagraphAlignment.Center

        Table3Column = Me.table3FirstPage.AddColumn("0.0001cm")
        Table3Column.Format.Alignment = ParagraphAlignment.Center

        'Create the header of the Third table
        Dim Table3Row As Row = table3FirstPage.AddRow()
        Table3Row.Height = 23
        Table3Row.HeadingFormat = False
        Table3Row.Format.Font.Size = 11
        Table3Row.Format.Alignment = ParagraphAlignment.Center
        '-----
        Table3Row.Cells(0).AddParagraph("Gauge Employed")
        Table3Row.Cells(0).Format.Alignment = ParagraphAlignment.Center
        Table3Row.Cells(0).VerticalAlignment = VerticalAlignment.Center
        Table3Row.Cells(0).Shading.Color = CellBackgroung_GBlue
        '-----
        Table3Row.Cells(1).AddParagraph("Serial No.")
        Table3Row.Cells(1).Format.Alignment = ParagraphAlignment.Center
        Table3Row.Cells(1).VerticalAlignment = VerticalAlignment.Center
        Table3Row.Cells(1).Shading.Color = CellBackgroung_GBlue
        '-----
        Table3Row.Cells(2).AddParagraph("Certification Date")
        Table3Row.Cells(2).Format.Alignment = ParagraphAlignment.Center
        Table3Row.Cells(2).VerticalAlignment = VerticalAlignment.Center
        Table3Row.Cells(2).Shading.Color = CellBackgroung_GBlue
        '-----
        Table3Row.Cells(3).AddParagraph("Next Certification Date")
        Table3Row.Cells(3).Format.Alignment = ParagraphAlignment.Center
        Table3Row.Cells(3).VerticalAlignment = VerticalAlignment.Center
        Table3Row.Cells(3).Shading.Color = CellBackgroung_GBlue
        '-----
        Table3Row.Cells(4).AddParagraph()
        '-----

        'Create a loop that will form the other boxes
        For c As Integer = 1 To 3
            Table3Row = table3FirstPage.AddRow()
            Table3Row.Height = 20.67
            For i As Integer = 0 To 4
                Table3Row.Cells(i).AddParagraph()
            Next
        Next
        '-----------END OF THIRD TABLE-------------

        'This is what's at the bottom of the first page

        paragraph = section.AddParagraph()

        'Add the text frame
        Me.BelowTable3 = section.AddTextFrame()
        Me.BelowTable3.Width = "20cm"
        Me.BelowTable3.Left = ShapePosition.Center
        Me.BelowTable3.RelativeHorizontal = RelativeHorizontal.Margin
        Me.BelowTable3.RelativeVertical = RelativeVertical.Page

        ' Writing what's below the third table
        paragraph = Me.BelowTable3.AddParagraph()
        For c As Integer = 1 To 51
            paragraph.AddLineBreak()
        Next
        paragraph.AddTab()
        paragraph.AddText("Comments: ________________________________________________________________________")
        paragraph.AddTab()
        paragraph.AddTab()
        paragraph.AddText("__________________________________________________________________________________")
        paragraph.Format.Font.Name = "Arial"
        paragraph.Format.Font.Size = 11
        paragraph.AddLineBreak()
        paragraph.AddLineBreak()
        paragraph.AddLineBreak()
        paragraph.AddTab()
        paragraph.AddTab()
        paragraph.AddText("Inspected by: ______________________________________")
        paragraph.AddTab()
        paragraph.AddText("Date: ________________")
        paragraph.AddLineBreak()
        paragraph.AddLineBreak()
        paragraph.AddLineBreak()
        paragraph.AddTab()
        paragraph.AddTab()
        paragraph.AddText("Review by: ________________________________________")
        paragraph.AddTab()
        paragraph.AddText("Date: ________________")
        paragraph = section.AddParagraph()
        Me.AboveTable3.Section.AddPageBreak()

        '**********************************
        '*--------END OF FIRST PAGE-------*
        '**********************************

        '*************************************
        '*--------START OF SECOND PAGE-------*
        '*************************************

        'Add the address text frame
        Me.AboveTable4 = section.AddTextFrame()
        Me.AboveTable4.Width = "5cm"
        Me.AboveTable4.Left = ShapePosition.Center
        Me.AboveTable4.RelativeHorizontal = RelativeHorizontal.Margin
        Me.AboveTable4.RelativeVertical = RelativeVertical.Page

        ' Writing what's above the fourth table
        paragraph = Me.AboveTable4.AddParagraph()
        paragraph.AddLineBreak()
        paragraph.AddLineBreak()
        paragraph.AddLineBreak()
        paragraph.AddLineBreak()
        paragraph.AddTab()
        paragraph.AddFormattedText("Variable Data", TextFormat.Bold)
        paragraph.Format.Font.Size = 12
        paragraph.AddLineBreak()
        For i As Integer = 0 To 5
            paragraph = section.AddParagraph()
        Next

        'create the fourth Table
        Me.table4SecondPage = section.AddTable()
        Me.table4SecondPage.Style = "Table"
        Me.table4SecondPage.Borders.Color = TableBorder_Black
        Me.table4SecondPage.Borders.Width = 0.25
        Me.table4SecondPage.Borders.Left.Width = 0.7
        Me.table4SecondPage.Borders.Right.Width = 0.7
        Me.table4SecondPage.Rows.LeftIndent = 15.1

        'Create the columns of the fourth table
        '17.941cm

        Dim Table4Column As Column = Me.table4SecondPage.AddColumn("1.993cm")
        Table4Column.Format.Alignment = ParagraphAlignment.Center

        Table4Column = Me.table4SecondPage.AddColumn("1.993cm")
        Table4Column.Format.Alignment = ParagraphAlignment.Center

        Table4Column = Me.table4SecondPage.AddColumn("1.397cm")
        Table4Column.Format.Alignment = ParagraphAlignment.Center

        Table4Column = Me.table4SecondPage.AddColumn("0.597cm")
        Table4Column.Format.Alignment = ParagraphAlignment.Center

        Table4Column = Me.table4SecondPage.AddColumn("1.993cm")
        Table4Column.Format.Alignment = ParagraphAlignment.Center

        Table4Column = Me.table4SecondPage.AddColumn("1.993cm")
        Table4Column.Format.Alignment = ParagraphAlignment.Center

        Table4Column = Me.table4SecondPage.AddColumn("1.993cm")
        Table4Column.Format.Alignment = ParagraphAlignment.Center

        Table4Column = Me.table4SecondPage.AddColumn("1.993cm")
        Table4Column.Format.Alignment = ParagraphAlignment.Center

        Table4Column = Me.table4SecondPage.AddColumn("1.993cm")
        Table4Column.Format.Alignment = ParagraphAlignment.Center

        Table4Column = Me.table4SecondPage.AddColumn("1.993cm")
        Table4Column.Format.Alignment = ParagraphAlignment.Center

        'Create the header of the Third table
        Dim Table4Row As Row = table4SecondPage.AddRow()
        Table4Row.HeadingFormat = False
        Table4Row.Format.Font.Size = 11
        Table4Row.Format.Alignment = ParagraphAlignment.Center
        Table4Row.Height = 15
        '-----
        Table4Row.Cells(0).AddParagraph("Care Fusion / Vendor Lot No.")
        Table4Row.Cells(0).MergeRight = 2
        Table4Row.Cells(0).VerticalAlignment = VerticalAlignment.Center
        Table4Row.Cells(0).Shading.Color = CellBackgroung_GBlue
        '-----
        Table4Row.Cells(3).AddParagraph()
        Table4Row.Cells(3).MergeRight = 2
        '-----
        Table4Row.Cells(6).AddParagraph("Part No.")
        Table4Row.Cells(6).VerticalAlignment = VerticalAlignment.Center
        Table4Row.Cells(6).Shading.Color = CellBackgroung_GBlue
        '-----
        Table4Row.Cells(7).AddParagraph()
        Table4Row.Cells(7).MergeRight = 2
        '---------

        Table4Row = table4SecondPage.AddRow()
        Table4Row.Height = 25
        Table4Row.Cells(0).Format.Font.Size = 11
        Table4Row.Cells(0).AddParagraph("Replicate")
        Table4Row.Cells(0).VerticalAlignment = VerticalAlignment.Center
        Table4Row.Cells(0).Shading.Color = CellBackgroung_GBlue

        'This will change the background color to blue in the second row of the header
        For i As Integer = 1 To 9
            If i = 2 Then
                Table4Row(i).MergeRight = 1
                Table4Row.Cells(i).Shading.Color = CellBackgroung_GBlue
            Else
                Table4Row.Cells(i).Shading.Color = CellBackgroung_GBlue
            End If
        Next

        'This loop will create the rest of the rows of the table
        For i As Integer = 0 To 49
            Table4Row = table4SecondPage.AddRow()
            Table4Row.Height = 12
            Table4Row.Cells(0).Format.Font.Size = 11
            Table4Row.Cells(0).AddParagraph(i + 1)
            Table4Row.Cells(2).MergeRight = 1
        Next

        Me.table4SecondPage.Section.AddPageBreak()

        '***********************************
        '*--------END OF SECOND PAGE-------*
        '***********************************

        '************************************
        '*--------START OF THIRD PAGE-------*
        '************************************

        'Add the address text frame
        Me.AboveTable5 = section.AddTextFrame()
        Me.AboveTable5.Width = "5cm"
        Me.AboveTable5.Left = ShapePosition.Center
        Me.AboveTable5.RelativeHorizontal = RelativeHorizontal.Margin
        Me.AboveTable5.RelativeVertical = RelativeVertical.Page

        ' Writing what's above the fifth table
        paragraph = Me.AboveTable5.AddParagraph()
        paragraph.AddLineBreak()
        paragraph.AddLineBreak()
        paragraph.AddLineBreak()
        paragraph.AddLineBreak()
        paragraph.AddTab()
        paragraph.AddFormattedText("Visual Data", TextFormat.Bold)
        paragraph.Format.Font.Size = 12
        paragraph.AddLineBreak()
        For i As Integer = 0 To 5
            paragraph = section.AddParagraph()
        Next

        'create the fifth Table
        Me.table5ThirdPage = section.AddTable()
        Me.table5ThirdPage.Style = "Table"
        Me.table5ThirdPage.Borders.Color = TableBorder_Black
        Me.table5ThirdPage.Borders.Width = 0.25
        Me.table5ThirdPage.Borders.Left.Width = 0.7
        Me.table5ThirdPage.Borders.Right.Width = 0.7
        Me.table5ThirdPage.Rows.LeftIndent = 15.1

        'Create the columns of the fifth table
        '17.941cm
        '0
        Dim Table5Column As Column = Me.table5ThirdPage.AddColumn("1.943cm")
        Table5Column.Format.Alignment = ParagraphAlignment.Center
        ''1
        Table5Column = Me.table5ThirdPage.AddColumn("2.543cm")
        Table5Column.Format.Alignment = ParagraphAlignment.Center
        '2
        Table5Column = Me.table5ThirdPage.AddColumn("0.843cm")
        Table5Column.Format.Alignment = ParagraphAlignment.Center
        '3
        Table5Column = Me.table5ThirdPage.AddColumn("1.100cm")
        Table5Column.Format.Alignment = ParagraphAlignment.Center
        ''4
        Table5Column = Me.table5ThirdPage.AddColumn("2.493cm")
        Table5Column.Format.Alignment = ParagraphAlignment.Center
        '5
        Table5Column = Me.table5ThirdPage.AddColumn("0.772cm")
        Table5Column.Format.Alignment = ParagraphAlignment.Center
        '6
        Table5Column = Me.table5ThirdPage.AddColumn("1.172cm")
        Table5Column.Format.Alignment = ParagraphAlignment.Center
        ''7
        Table5Column = Me.table5ThirdPage.AddColumn("0.571cm")
        Table5Column.Format.Alignment = ParagraphAlignment.Center
        ''8
        Table5Column = Me.table5ThirdPage.AddColumn("1.972cm")
        Table5Column.Format.Alignment = ParagraphAlignment.Center
        '9
        Table5Column = Me.table5ThirdPage.AddColumn("1.943cm")
        Table5Column.Format.Alignment = ParagraphAlignment.Center
        ''10
        Table5Column = Me.table5ThirdPage.AddColumn("2.543cm")
        Table5Column.Format.Alignment = ParagraphAlignment.Center

        'Create the header of the fifth table
        Dim Table5Row As Row = table5ThirdPage.AddRow()
        Table5Row.HeadingFormat = False
        Table5Row.Format.Font.Size = 11
        Table5Row.Format.Alignment = ParagraphAlignment.Center
        Table5Row.Height = 16
        '-----
        Table5Row.Cells(0).AddParagraph("Care Fusion / Vendor Lot No.")
        Table5Row.Cells(0).MergeRight = 2
        Table5Row.Cells(0).VerticalAlignment = VerticalAlignment.Center
        Table5Row.Cells(0).Shading.Color = CellBackgroung_GBlue
        '-----
        Table5Row.Cells(3).MergeRight = 2
        '-----
        Table5Row.Cells(6).AddParagraph("Part No.")
        Table5Row.Cells(6).MergeRight = 1
        Table5Row.Cells(6).VerticalAlignment = VerticalAlignment.Center
        Table5Row.Cells(6).Shading.Color = CellBackgroung_GBlue
        '-----
        Table5Row.Cells(8).MergeRight = 2

        '---------- Second row of the header
        Table5Row = table5ThirdPage.AddRow()
        Table5Row.Format.Font.Size = 11
        Table5Row.Height = 15
        Table5Row.Cells(2).MergeRight = 1
        '-----
        Table5Row.Cells(5).MergeRight = 1
        '-----
        Table5Row.Cells(7).AddParagraph()
        Table5Row.Cells(7).MergeRight = 1
        For i As Integer = 0 To 10
            If i = 0 Or i = 2 Or i = 5 Or i = 9 Then
                Table5Row.Cells(i).AddParagraph("Replicate")
                Table5Row.Cells(i).Format.Alignment = ParagraphAlignment.Center
                Table5Row.Cells(i).VerticalAlignment = VerticalAlignment.Center
            End If
            Table5Row.Cells(i).Shading.Color = CellBackgroung_GBlue
        Next

        'The code and loop that will create the rest of the table
        For i As Integer = 0 To 49
            Table5Row = table5ThirdPage.AddRow()
            Table5Row.Format.Font.Size = 11
            Table5Row.Format.Alignment = ParagraphAlignment.Center
            Table5Row.VerticalAlignment = VerticalAlignment.Center

            For c As Integer = 0 To 10
                If c = 2 Or c = 5 Or c = 7 Then
                    Table5Row.Cells(c).MergeRight = 1
                End If
                If c = 1 Or c = 4 Or c = 7 Or c = 10 Then
                    Table5Row.Cells(c).AddParagraph(" pass  fail")
                End If
            Next
            Table5Row.Cells(0).AddParagraph(i + 1)
            Table5Row.Cells(2).AddParagraph(i + 51)
            Table5Row.Cells(5).AddParagraph(i + 101)
            Table5Row.Cells(9).AddParagraph(i + 151)
        Next

        'This is what's at the bottom of the third page

        paragraph = section.AddParagraph()

        'Add the text frame
        Me.BelowTable5 = section.AddTextFrame()
        Me.BelowTable5.Width = "20cm"
        Me.BelowTable5.Left = ShapePosition.Center
        Me.BelowTable5.RelativeHorizontal = RelativeHorizontal.Margin
        Me.BelowTable5.RelativeVertical = RelativeVertical.Page

        ' Writing what's below the fifth table
        paragraph = Me.BelowTable5.AddParagraph()
        paragraph.Format.Font.Size = 9
        For c As Integer = 1 To 73
            paragraph.AddLineBreak()
        Next
        paragraph.AddTab()
        paragraph.AddText("Note: In the event that different visual defect classification inspections are required, print this page as need. ")


    End Sub

End Class

'End Namespace

