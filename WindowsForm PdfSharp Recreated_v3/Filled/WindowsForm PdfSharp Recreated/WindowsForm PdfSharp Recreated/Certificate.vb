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

    'Will store all the information about what is written
    Public Foot As Footer
    Public Head As Header
    Public Body As Document_info
    Public Sub New(ByRef footer As Footer, ByRef header As Header, ByRef infor As Document_info)
        Foot = footer
        Head = header
        Body = infor
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
        section.PageSetup.BottomMargin = "1.6cm"
        section.PageSetup.TopMargin = "2.8cm"
        section.PageSetup.RightMargin = "1cm"
        section.PageSetup.LeftMargin = "1cm"

        'Create the header
        PicL = section.Headers.Primary.AddImage("C:\Users\Kevin\Documents\Visual Studio 2015\Projects\WindowsForm PdfSharp Recreated\WindowsForm PdfSharp Recreated\Logo_Bard.png")
        PicL.Height = "1.7cm"
        PicL.LockAspectRatio = True
        PicL.RelativeVertical = RelativeVertical.Page
        PicL.RelativeHorizontal = RelativeHorizontal.Margin
        PicL.Top = "0.25cm"
        PicL.Left = ShapePosition.Left
        PicL.WrapFormat.Style = WrapStyle.Through
        'section.Headers.Primary.Format.Borders.Bottom.Visible = True

        Dim paragraph As Paragraph = section.Headers.Primary.AddParagraph
        paragraph.AddFormattedText(Head.Reference, TextFormat.Italic)
        paragraph.AddLineBreak()
        paragraph.Format.Font.Size = 9
        paragraph.Format.Alignment = ParagraphAlignment.Right

        'Create the footer
        paragraph = section.Footers.Primary.AddParagraph
        paragraph.Format.Font.Name = "Times New Roman"
        paragraph.AddText(Foot.FormNumber)
        paragraph.AddSpace(43)
        paragraph.AddText(Foot.RevJDate)
        paragraph.AddTab()
        paragraph.AddTab()
        paragraph.AddTab()
        paragraph.AddText(Foot.CAFNumberDate)
        paragraph.AddTab()
        paragraph.AddText("Page ")
        paragraph.AddPageField()
        paragraph.AddText(" of ")
        paragraph.AddNumPagesField()
        paragraph.Format.Font.Size = 10
        paragraph.Format.Alignment = ParagraphAlignment.Center
        'section.Footers.Primary.Format.Borders.Top.Visible = True


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
        paragraph = section.AddParagraph()

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
        row.Cells(1).Format.Font.Size = 10
        row.Cells(1).AddParagraph(Body.Table1.Supplier)
        row.Cells(1).Format.Alignment = ParagraphAlignment.Center
        row.Cells(1).VerticalAlignment = VerticalAlignment.Center
        '-----
        row.Cells(2).Format.Font.Size = 11
        row.Cells(2).AddParagraph("Data Recieved")
        row.Cells(2).Format.Alignment = ParagraphAlignment.Right
        row.Cells(2).VerticalAlignment = VerticalAlignment.Center
        row.Cells(2).Shading.Color = CellBackgroung_GBlue
        '-----
        row.Cells(3).Format.Font.Size = 10
        row.Cells(3).AddParagraph(Body.Table1.Data_Recieved)
        row.Cells(3).Format.Alignment = ParagraphAlignment.Center
        row.Cells(3).VerticalAlignment = VerticalAlignment.Center

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
        row.Cells(1).Format.Font.Size = 10
        row.Cells(1).AddParagraph(Body.Table1.PO_Number)
        row.Cells(1).Format.Alignment = ParagraphAlignment.Center
        row.Cells(1).VerticalAlignment = VerticalAlignment.Center
        '-----
        row.Cells(2).Format.Font.Size = 11
        row.Cells(2).AddParagraph("Reciever No")
        row.Cells(2).Format.Alignment = ParagraphAlignment.Right
        row.Cells(2).VerticalAlignment = VerticalAlignment.Center
        row.Cells(2).Shading.Color = CellBackgroung_GBlue
        '-----
        row.Cells(3).Format.Font.Size = 10
        row.Cells(3).AddParagraph(Body.Table1.Reciever_Number)
        row.Cells(3).Format.Alignment = ParagraphAlignment.Center
        row.Cells(3).VerticalAlignment = VerticalAlignment.Center


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
        row.Cells(1).Format.Font.Size = 10
        row.Cells(1).AddParagraph(Body.Table1.ManuExpi_Date)
        row.Cells(1).Format.Alignment = ParagraphAlignment.Center
        row.Cells(1).VerticalAlignment = VerticalAlignment.Center
        '-----
        row.Cells(2).Format.Font.Size = 11
        row.Cells(2).AddParagraph("CareFusion /")
        row.Cells(2).AddParagraph(" Vendor Lot No /")
        row.Cells(2).Format.Alignment = ParagraphAlignment.Right
        row.Cells(2).VerticalAlignment = VerticalAlignment.Center
        row.Cells(2).Shading.Color = CellBackgroung_GBlue
        '-----
        row.Cells(3).Format.Font.Size = 10
        row.Cells(3).AddParagraph(Body.Table1.CareVendor_Number)
        row.Cells(3).Format.Alignment = ParagraphAlignment.Center
        row.Cells(3).VerticalAlignment = VerticalAlignment.Center


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
        row.Cells(1).Format.Font.Size = 10
        row.Cells(1).AddParagraph(Body.Table1.Part_Number)
        row.Cells(1).Format.Alignment = ParagraphAlignment.Center
        row.Cells(1).VerticalAlignment = VerticalAlignment.Center
        '-----
        row.Cells(2).Format.Font.Size = 11
        row.Cells(2).AddParagraph("RMS Revision")
        row.Cells(2).Format.Alignment = ParagraphAlignment.Right
        row.Cells(2).VerticalAlignment = VerticalAlignment.Center
        row.Cells(2).Shading.Color = CellBackgroung_GBlue
        '-----
        row.Cells(3).Format.Font.Size = 10
        row.Cells(3).AddParagraph(Body.Table1.RMS_Revision)
        row.Cells(3).Format.Alignment = ParagraphAlignment.Center
        row.Cells(3).VerticalAlignment = VerticalAlignment.Center

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
        row.Cells(1).Format.Font.Size = 10
        row.Cells(1).AddParagraph(Body.Table1.Quantity)
        row.Cells(1).Format.Alignment = ParagraphAlignment.Center
        row.Cells(1).VerticalAlignment = VerticalAlignment.Center
        '-----
        row.Cells(2).Format.Font.Size = 11
        row.Cells(2).AddParagraph("Containers Recieved")
        row.Cells(2).Format.Alignment = ParagraphAlignment.Right
        row.Cells(2).VerticalAlignment = VerticalAlignment.Center
        row.Cells(2).Shading.Color = CellBackgroung_GBlue
        '-----
        row.Cells(3).Format.Font.Size = 10
        row.Cells(3).AddParagraph(Body.Table1.Containers_Recieved)
        row.Cells(3).Format.Alignment = ParagraphAlignment.Center
        row.Cells(3).VerticalAlignment = VerticalAlignment.Center
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
        row.Cells(1).Format.Font.Size = 10
        row.Cells(1).AddParagraph(Body.Table1.Unit_of_Measure)
        row.Cells(1).Format.Alignment = ParagraphAlignment.Center
        row.Cells(1).VerticalAlignment = VerticalAlignment.Center
        '-----
        row.Cells(2).Format.Font.Size = 11
        row.Cells(2).AddParagraph("Boxes or Roll #")
        row.Cells(2).AddParagraph("inspected")
        row.Cells(2).Format.Alignment = ParagraphAlignment.Right
        row.Cells(2).VerticalAlignment = VerticalAlignment.Center
        row.Cells(2).Shading.Color = CellBackgroung_GBlue
        '-----
        row.Cells(3).Format.Font.Size = 10
        row.Cells(3).AddParagraph(Body.Table1.BoxRoll_Number)
        row.Cells(3).Format.Alignment = ParagraphAlignment.Center
        row.Cells(3).VerticalAlignment = VerticalAlignment.Center

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
        row.Cells(1).Format.Font.Size = 10
        row.Cells(1).AddParagraph(Body.Table1.Resin_Number)
        row.Cells(1).Format.Alignment = ParagraphAlignment.Center
        row.Cells(1).VerticalAlignment = VerticalAlignment.Center
        '-----
        row.Cells(2).Format.Font.Size = 11
        row.Cells(2).AddParagraph("Inspection Level")
        row.Cells(2).Format.Alignment = ParagraphAlignment.Right
        row.Cells(2).VerticalAlignment = VerticalAlignment.Center
        row.Cells(2).Shading.Color = CellBackgroung_GBlue
        '-----
        row.Cells(3).Format.Font.Size = 11
        row.Cells(3).AddParagraph(Body.Table1.Inspection_Level)
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
        row.Cells(1).Format.Font.Size = 10
        row.Cells(1).AddParagraph(Body.Table1.Sampled_By)
        row.Cells(1).Format.Alignment = ParagraphAlignment.Center
        row.Cells(1).VerticalAlignment = VerticalAlignment.Center
        '-----
        row.Cells(2).Format.Font.Size = 11
        row.Cells(2).AddParagraph("Data Sampled")
        row.Cells(2).Format.Alignment = ParagraphAlignment.Right
        row.Cells(2).VerticalAlignment = VerticalAlignment.Center
        row.Cells(2).Shading.Color = CellBackgroung_GBlue
        '-----
        row.Cells(3).Format.Font.Size = 10
        row.Cells(3).AddParagraph(Body.Table1.Data_Sampled)
        row.Cells(3).Format.Alignment = ParagraphAlignment.Center
        row.Cells(3).VerticalAlignment = VerticalAlignment.Center

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
        Table2Row.Format.Alignment = ParagraphAlignment.Center
        Table2Row.VerticalAlignment = VerticalAlignment.Center
        Table2Row.Cells(0).AddParagraph("Class II")
        Table2Row.Cells(0).AddParagraph("(Variable)")
        Table2Row.Cells(0).Format.Font.Size = 11
        Table2Row.Cells(0).Shading.Color = CellBackgroung_GBlue
        Table2Row.Cells(1).AddParagraph(Body.Table2.Sampling_Level(0))
        Table2Row.Cells(2).AddParagraph(Body.Table2.AQL(0))
        Table2Row.Cells(3).AddParagraph(Body.Table2.Sampling_Size(0))
        Table2Row.Cells(4).Format.Font.Size = 10
        Table2Row.Cells(4).AddParagraph("a = " & Body.Table2.Accept(0) & "   r = " & Body.Table2.Reject(0) & " ")
        Table2Row.Cells(5).AddParagraph(Body.Table2.Number_Of_Defects(0))
        Table2Row.Cells(6).AddParagraph(Body.Table2.Type_Of_Defect(0))
        Table2Row.Cells(7).AddParagraph(Body.Table2.Pass_Fail(0))

        '------ Second Row of Data -----
        Table2Row = table2FirstPage.AddRow()
        Table2Row.Height = 32
        Table2Row.Format.Alignment = ParagraphAlignment.Center
        Table2Row.VerticalAlignment = VerticalAlignment.Center
        Table2Row.Cells(0).AddParagraph("Class II")
        Table2Row.Cells(0).AddParagraph("(Visual)")
        Table2Row.Cells(0).Format.Font.Size = 11
        Table2Row.Cells(0).Shading.Color = CellBackgroung_GBlue
        Table2Row.Cells(1).AddParagraph(Body.Table2.Sampling_Level(1))
        Table2Row.Cells(2).AddParagraph(Body.Table2.AQL(1))
        Table2Row.Cells(3).AddParagraph(Body.Table2.Sampling_Size(1))
        Table2Row.Cells(4).Format.Font.Size = 10
        Table2Row.Cells(4).AddParagraph("a = " & Body.Table2.Accept(1) & "   r = " & Body.Table2.Reject(1) & " ")
        Table2Row.Cells(5).AddParagraph(Body.Table2.Number_Of_Defects(1))
        Table2Row.Cells(6).AddParagraph(Body.Table2.Type_Of_Defect(1))
        Table2Row.Cells(7).AddParagraph(Body.Table2.Pass_Fail(1))

        '------ Third Row of Data -----
        Table2Row = table2FirstPage.AddRow()
        Table2Row.Height = 32
        Table2Row.Format.Alignment = ParagraphAlignment.Center
        Table2Row.VerticalAlignment = VerticalAlignment.Center
        Table2Row.Cells(0).AddParagraph("Class III")
        Table2Row.Cells(0).AddParagraph("(Variable)")
        Table2Row.Cells(0).Format.Font.Size = 11
        Table2Row.Cells(0).Shading.Color = CellBackgroung_GBlue
        Table2Row.Cells(1).AddParagraph(Body.Table2.Sampling_Level(2))
        Table2Row.Cells(2).AddParagraph(Body.Table2.AQL(2))
        Table2Row.Cells(3).AddParagraph(Body.Table2.Sampling_Size(2))
        Table2Row.Cells(4).Format.Font.Size = 10
        Table2Row.Cells(4).AddParagraph("a = " & Body.Table2.Accept(2) & "   r = " & Body.Table2.Reject(2) & " ")
        Table2Row.Cells(5).AddParagraph(Body.Table2.Number_Of_Defects(2))
        Table2Row.Cells(6).AddParagraph(Body.Table2.Type_Of_Defect(2))
        Table2Row.Cells(7).AddParagraph(Body.Table2.Pass_Fail(2))

        '------ Fourth Row of Data -----
        Table2Row = table2FirstPage.AddRow()
        Table2Row.Height = 32
        Table2Row.Format.Alignment = ParagraphAlignment.Center
        Table2Row.VerticalAlignment = VerticalAlignment.Center
        Table2Row.Cells(0).AddParagraph("Class III")
        Table2Row.Cells(0).AddParagraph("(Visual)")
        Table2Row.Cells(0).Format.Font.Size = 11
        Table2Row.Cells(0).Shading.Color = CellBackgroung_GBlue
        Table2Row.Cells(1).AddParagraph(Body.Table2.Sampling_Level(3))
        Table2Row.Cells(2).AddParagraph(Body.Table2.AQL(3))
        Table2Row.Cells(3).AddParagraph(Body.Table2.Sampling_Size(3))
        Table2Row.Cells(4).Format.Font.Size = 10
        Table2Row.Cells(4).AddParagraph("a = " & Body.Table2.Accept(3) & "   r = " & Body.Table2.Reject(3) & " ")
        Table2Row.Cells(5).AddParagraph(Body.Table2.Number_Of_Defects(3))
        Table2Row.Cells(6).AddParagraph(Body.Table2.Type_Of_Defect(3))
        Table2Row.Cells(7).AddParagraph(Body.Table2.Pass_Fail(3))

        '------ Fifth Row of Data -----
        Table2Row = table2FirstPage.AddRow()
        Table2Row.Height = 32
        Table2Row.Format.Alignment = ParagraphAlignment.Center
        Table2Row.VerticalAlignment = VerticalAlignment.Center
        Table2Row.Cells(0).AddParagraph("Class IV")
        Table2Row.Cells(0).AddParagraph("(Visual)")
        Table2Row.Cells(0).Format.Font.Size = 11
        Table2Row.Cells(0).Shading.Color = CellBackgroung_GBlue
        Table2Row.Cells(1).AddParagraph(Body.Table2.Sampling_Level(4))
        Table2Row.Cells(2).AddParagraph(Body.Table2.AQL(4))
        Table2Row.Cells(3).AddParagraph(Body.Table2.Sampling_Size(4))
        Table2Row.Cells(4).Format.Font.Size = 10
        Table2Row.Cells(4).AddParagraph("a = " & Body.Table2.Accept(4) & "   r = " & Body.Table2.Reject(4) & " ")
        Table2Row.Cells(5).AddParagraph(Body.Table2.Number_Of_Defects(4))
        Table2Row.Cells(6).AddParagraph(Body.Table2.Type_Of_Defect(4))
        Table2Row.Cells(7).AddParagraph(Body.Table2.Pass_Fail(4))

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
        If Body.Non_Table_Info.Sample_Amount > 0 Then
            paragraph.AddFormattedText("R", New Font("Wingdings 2"))
            paragraph.AddText(" Lot # inspected from previous shipment sample size increased by ")
            paragraph.AddFormattedText("    " & Body.Non_Table_Info.Sample_Amount & "    ", TextFormat.Underline)
            paragraph.AddText(" samples.  ")
        Else
            paragraph.AddFormattedText("c", New Font("Webdings"))
            paragraph.AddFormattedText(" Lot # inspected from previous shipment sample size increased by __________ samples.  ")
        End If
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


        'This is meant as an example of how the "AddData" function works
        Body.Table3.AddData("Bard", "2341182", "7/29/2016", "8/29/2016")
        Body.Table3.AddData("P-Zero", "8920842", "7/29/2016", "8/29/2016")
        Body.Table3.AddData("3D-LMI", "4475832", "7/29/2016", "8/29/2016")

        'The following code will create and fill the rest of the rows of the table
        Dim Values_T3 As IEnumerable(Of Document_info.Third_Table)
        Values_T3 = Body.Table3.GetData()
        If Body.Table3.Is_Empty = True Then 'If the table is empty, only create one empty row
            Table3Row = table3FirstPage.AddRow()
            Table3Row.Height = 20.67
        Else
            For Each da As Document_info.Third_Table In Values_T3
                Table3Row = table3FirstPage.AddRow()
                Table3Row.Height = 20.67
                For r As Integer = 0 To 3
                    Table3Row.Cells(r).Format.Alignment = ParagraphAlignment.Center
                    Table3Row.Cells(r).VerticalAlignment = VerticalAlignment.Center
                    If r = 0 Then
                        Table3Row.Cells(r).AddParagraph(da.Gauge_Employed)
                    ElseIf r = 1 Then
                        Table3Row.Cells(r).AddParagraph(da.Serial_Number)
                    ElseIf r = 2 Then
                        Table3Row.Cells(r).AddParagraph(da.Certification_Date)
                    Else
                        Table3Row.Cells(r).AddParagraph(da.Next_Certification_Date)
                    End If
                Next
            Next
        End If

        '-----------END OF THIRD TABLE-------------

        'This is what's at the bottom of the first page

        paragraph = section.AddParagraph()

        'Add the text frame
        Me.BelowTable3 = section.AddTextFrame()
        Me.BelowTable3.Width = "20cm"
        Me.BelowTable3.Left = ShapePosition.Center
        Me.BelowTable3.RelativeHorizontal = RelativeHorizontal.Margin
        Me.BelowTable3.RelativeVertical = RelativeVertical.Paragraph

        ' Writing what's below the third table
        paragraph = Me.BelowTable3.AddParagraph()
        paragraph.AddTab()
        paragraph.Format.Font.Size = 11
        paragraph.Format.Font.Name = "Arial"
        If Body.Non_Table_Info.Comment <> "" Then
            paragraph.AddText("Comments: ")
            If Body.Non_Table_Info.Comment.Length > 86 Then
                Dim First_Comment As String = Body.Non_Table_Info.Comment.Remove(86)
                Dim Second_Comment As String = Body.Non_Table_Info.Comment.Remove(0, 86)
                paragraph.AddFormattedText(First_Comment, TextFormat.Underline)
                paragraph.AddLineBreak()
                paragraph.AddTab()
                paragraph.AddFormattedText(Second_Comment, TextFormat.Underline)
            Else
                paragraph.AddFormattedText(Body.Non_Table_Info.Comment, TextFormat.Underline)
            End If
        Else
            paragraph.AddText("Comments: N/A")
        End If
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
        Me.BelowTable3.Section.AddPageBreak()

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
        Table4Row.Cells(3).AddParagraph(Body.Table4.CareVendorLot_Number)
        Table4Row.Cells(3).MergeRight = 2
        '-----
        Table4Row.Cells(6).AddParagraph("Part No.")
        Table4Row.Cells(6).VerticalAlignment = VerticalAlignment.Center
        Table4Row.Cells(6).Shading.Color = CellBackgroung_GBlue
        '-----
        Table4Row.Cells(7).AddParagraph(Body.Table4.Part_Number)
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
            Table4Row.Cells(i).Format.Font.Size = 11
            Table4Row.Cells(i).VerticalAlignment = VerticalAlignment.Center
            Table4Row.Cells(i).Format.Alignment = ParagraphAlignment.Center
            Table4Row.Cells(i).Shading.Color = CellBackgroung_GBlue
            If i = 1 Then
                Table4Row.Cells(i).AddParagraph("Filler " & i)
            ElseIf i = 2 Then
                Table4Row(i).MergeRight = 1
                Table4Row.Cells(i).AddParagraph("Filler " & i)
            ElseIf i = 3 Then
                Table4Row.Cells(i).AddParagraph("Filler " & i)
            ElseIf i = 4 Then
                Table4Row.Cells(i).AddParagraph("Filler " & i)
            ElseIf i = 5 Then
                Table4Row.Cells(i).AddParagraph("Filler " & i)
            ElseIf i = 6 Then
                Table4Row.Cells(i).AddParagraph("Filler " & i)
            ElseIf i = 7 Then
                Table4Row.Cells(i).AddParagraph("Filler " & i)
            ElseIf i = 8 Then
                Table4Row.Cells(i).AddParagraph("Filler " & i)
            ElseIf i = 9 Then
                Table4Row.Cells(i).AddParagraph("Filler " & i)
            End If
        Next

        'This loop is meant as an example of using the function "AddData" to add information to the table
        For d As Integer = 1 To 56
            Body.Table4.AddData("R: " & d & " C: 1", "R: " & d & " C: 2", "R: " & d & " C: 3", "R: " & d & " C: 4",
           "R: " & d & " C: 5", "R: " & d & " C: 6", "R: " & d & " C: 7", "R: " & d & " C: 8", "R: " & d & " C: 9")
        Next

        'The following code will create and fill the rest of the rows of the table

        Dim Values_T4 As IEnumerable(Of Document_info.Fourth_Table)
        If Body.Table4.Is_Empty = True Then 'If the table is empty, only create one empty row
            Table4Row = table4SecondPage.AddRow()
            Table4Row.Height = 12
            Table4Row.Cells(2).MergeRight = 1
        Else
            Dim number As Integer = 1
            Values_T4 = Body.Table4.GetData()
            For Each da As Document_info.Fourth_Table In Values_T4
                Table4Row = table4SecondPage.AddRow()
                Table4Row.Height = 12

                For r As Integer = 0 To 9
                    Table4Row.Cells(r).Format.Font.Size = 11
                    Table4Row.Cells(r).Format.Alignment = ParagraphAlignment.Left
                    Table4Row.Cells(r).VerticalAlignment = VerticalAlignment.Center
                    Table4Row.Cells(r).Format.Font.Size = 9
                    If r = 0 Then
                        Table4Row.Cells(r).Format.Alignment = ParagraphAlignment.Center
                        Table4Row.Cells(r).Format.Font.Size = 11
                        Table4Row.Cells(r).AddParagraph(number)
                    ElseIf r = 1 Then
                        Table4Row.Cells(r).AddParagraph(da.Filler_1)
                    ElseIf r = 2 Then
                        Table4Row.Cells(r).MergeRight = 1
                        Table4Row.Cells(r).AddParagraph(da.Filler_2)
                    ElseIf r = 3 Then
                        Table4Row.Cells(r).AddParagraph(da.Filler_3)
                    ElseIf r = 4 Then
                        Table4Row.Cells(r).AddParagraph(da.Filler_4)
                    ElseIf r = 5 Then
                        Table4Row.Cells(r).AddParagraph(da.Filler_5)
                    ElseIf r = 6 Then
                        Table4Row.Cells(r).AddParagraph(da.Filler_6)
                    ElseIf r = 7 Then
                        Table4Row.Cells(r).AddParagraph(da.Filler_7)
                    ElseIf r = 8 Then
                        Table4Row.Cells(r).AddParagraph(da.Filler_8)
                    ElseIf r = 9 Then
                        Table4Row.Cells(r).AddParagraph(da.Filler_9)
                    End If
                Next
                number += 1
            Next
        End If
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
        Table5Row.Cells(3).AddParagraph(Body.Table5.CareVendorLot_Number)
        Table5Row.Cells(3).MergeRight = 2
        Table5Row.Cells(3).VerticalAlignment = VerticalAlignment.Center
        '-----
        Table5Row.Cells(6).AddParagraph("Part No.")
        Table5Row.Cells(6).MergeRight = 1
        Table5Row.Cells(6).VerticalAlignment = VerticalAlignment.Center
        Table5Row.Cells(6).Shading.Color = CellBackgroung_GBlue
        '-----
        Table5Row.Cells(8).AddParagraph(Body.Table5.Part_Number)
        Table5Row.Cells(8).MergeRight = 2
        Table5Row.Cells(8).VerticalAlignment = VerticalAlignment.Center

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
        Body.Table5.AddData(True)
        Body.Table5.AddData(False)
        Body.Table5.AddData(True)
        Body.Table5.AddData(False)

        Dim answer() As String = Body.Table5.GetData()
        For i As Integer = 0 To 49
            Table5Row = table5ThirdPage.AddRow()
            Table5Row.Format.Font.Size = 11
            Table5Row.Format.Alignment = ParagraphAlignment.Center
            Table5Row.VerticalAlignment = VerticalAlignment.Center

            For c As Integer = 0 To 10
                If c = 2 Or c = 5 Or c = 7 Then
                    Table5Row.Cells(c).MergeRight = 1
                End If
            Next
            Table5Row.Cells(0).AddParagraph(i + 1)
            Table5Row.Cells(1).AddParagraph(answer(i))
            Table5Row.Cells(2).AddParagraph(i + 51)
            Table5Row.Cells(4).AddParagraph(answer(i + 50))
            Table5Row.Cells(5).AddParagraph(i + 101)
            Table5Row.Cells(7).AddParagraph(answer(i + 100))
            Table5Row.Cells(9).AddParagraph(i + 151)
            Table5Row.Cells(10).AddParagraph(answer(i + 150))
        Next

        'This is what's at the bottom of the third page
        Table5Row = table5ThirdPage.AddRow()
        Table5Row.Format.Alignment = ParagraphAlignment.Left
        Table5Row.Cells(0).MergeRight = 10
        Table5Row.Cells(0).Format.Font.Size = 9
        Table5Row.Cells(0).AddParagraph("Note: In the event that different visual defect classification inspections are required, print this page as need. ")
        Table5Row.Cells(0).VerticalAlignment = VerticalAlignment.Top
        Table5Row.Borders.Right.Visible = False
        Table5Row.Borders.Left.Visible = False
        Table5Row.Borders.Bottom.Clear()

    End Sub

End Class

'End Namespace