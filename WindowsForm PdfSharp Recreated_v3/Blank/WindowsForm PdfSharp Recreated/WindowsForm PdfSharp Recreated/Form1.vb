
Imports MigraDoc.DocumentObjectModel
Imports MigraDoc.Rendering

Public Class Form1
    Private Sub btnGenerateCertificate_Click(sender As Object, e As EventArgs) Handles btnGenerateCertificate.Click
        Dim regDate As Date = Date.Now()
        Dim strDate As String = regDate.ToString("MMddyyyy") & "_" & regDate.ToString("HHmm")
        'Variables fot path and filename 
        Dim filename As String
        Dim folderBrowser As New FolderBrowserDialog
        Dim Foot As Footer = New Footer
        Dim Head As Header = New Header
        Dim Body As Document_info = New Document_info

        Body.Table1.Supplier = Supplier.Text
        Body.Table1.Data_Recieved = Data_Recieved.Text
        Body.Table1.PO_Number = PO_Number.Text

        Try

            filename = "ReporteVal_Carefusion_" & strDate & ".pdf"

            folderBrowser.SelectedPath = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            folderBrowser.Description = "Select a destination folder for the Certificate of Compliance."
            If Not folderBrowser.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                MsgBox("Error, a folder must be selected to save the Certification.", MsgBoxStyle.Exclamation, "Test Report")
                Exit Sub
            End If

            MsgBox("EL folder seleccionado es:" & folderBrowser.SelectedPath)

            'Create a object certificate
            Dim certificate As Certificate = New Certificate(Foot, Head, Body)

            Dim document As Document = certificate.CreateDocument()
            document.UseCmykColor = True
            'Const unicode As Boolean = False

            'Const embedding As PdfFontEmbedding = PdfFontEmbedding.Always

            Dim pdfRenderer As PdfDocumentRenderer = New PdfDocumentRenderer(True)

            pdfRenderer.Document = document

            pdfRenderer.RenderDocument()


            pdfRenderer.PdfDocument.Save(folderBrowser.SelectedPath & "\" & filename)
            MsgBox("The Certified was saved as: " & folderBrowser.SelectedPath & "\" & filename)
            Process.Start(folderBrowser.SelectedPath & "\" & filename)
        Catch ex As Exception
            MsgBox("Error Migradosc: " & vbCrLf & ex.Message, MsgBoxStyle.Critical)
        End Try

    End Sub

End Class
