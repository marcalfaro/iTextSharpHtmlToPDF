Option Explicit On

Imports System.IO
Imports iTextSharp.text
Imports iTextSharp.text.html.simpleparser
Imports iTextSharp.text.pdf
Imports iTextSharp.tool.xml
Imports iTextSharp.tool.xml.html
Imports iTextSharp.tool.xml.parser
Imports iTextSharp.tool.xml.pipeline.css
Imports iTextSharp.tool.xml.pipeline.end
Imports iTextSharp.tool.xml.pipeline.html

Public Class Form1


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim htm As String = File.ReadAllText(Path.Combine(Application.StartupPath, "table2.htm"))

        With wb
            .ScriptErrorsSuppressed = True
            .Navigate("about:blank")  'needed to initialize webbrowser
            If .Document IsNot Nothing Then
                .Document.Write(htm)
            Else
                .DocumentText = htm
            End If
        End With

        Dim pdfFile As String = Path.Combine(Application.StartupPath, "table2.pdf")
        CreatePDFFromHTMLFile(htm, pdfFile)
    End Sub

    Public Sub CreatePDFFromHTMLFile(ByVal htmlText As String, ByVal TargetFile As String)
#Region " Example 1 "
        'Dim doc As New Document
        'Try
        '    PdfWriter.GetInstance(doc, New FileStream(TargetFile, FileMode.Create))
        '    doc.Open()

        '    Dim htmlarraylist As List(Of IElement) = HTMLWorker.ParseToList(New StringReader(htmlText), Nothing)
        '    For k As Integer = 0 To htmlarraylist.Count - 1
        '        doc.Add((htmlarraylist(k)))
        '    Next
        '    doc.Close()

        'Catch ex As Exception
        '    Debug.Print(ex.Message)
        'End Try
#End Region

#Region " Example 2 "
        'Try
        '    Using doc As New Document(PageSize.A4, 10.0F, 10.0F, 10.0F, 0F)
        '        PdfWriter.GetInstance(doc, New FileStream(TargetFile, FileMode.Create))

        '        doc.Open()

        '        Dim htmlarraylist As List(Of IElement) = HTMLWorker.ParseToList(New StringReader(htmlText), Nothing)
        '        For k As Integer = 0 To htmlarraylist.Count - 1
        '            doc.Add((htmlarraylist(k)))
        '        Next
        '        doc.Close()

        '    End Using

        'Catch ex As Exception
        '    Debug.Print(ex.Message)
        'End Try
#End Region

#Region " Example 3 "
        'Try
        '    Using doc As New Document(PageSize.A4, 10.0F, 10.0F, 10.0F, 0F)

        '        Dim rslt As Byte() = {}

        '        'Dim cssFiles As New List(Of String)
        '        'cssFiles.Add(Path.Combine(Application.StartupPath, "css1.css"))

        '        Using ms As New MemoryStream

        '            Dim writer As PdfWriter = PdfWriter.GetInstance(doc, New FileStream(TargetFile, FileMode.Create))
        '            'writer.CloseStream = False

        '            doc.Open()
        '            'doc.Add("")

        '            Dim htmlContext As HtmlPipelineContext = New HtmlPipelineContext(Nothing)
        '            htmlContext.SetTagFactory(Tags.GetHtmlTagProcessorFactory())

        '            Dim cssResolver As ICSSResolver = XMLWorkerHelper.GetInstance().GetDefaultCssResolver(False)
        '            cssResolver.AddCssFile(Path.Combine(Application.StartupPath, "table2.css"), True)

        '            Dim pipeline As IPipeline = New CssResolverPipeline(cssResolver, New HtmlPipeline(htmlContext, New PdfWriterPipeline(doc, writer)))
        '            Dim worker As XMLWorker = New XMLWorker(pipeline, True)
        '            Dim xparser As XMLParser = New XMLParser(worker)
        '            xparser.Parse(New MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlText)))
        '            'xparser.Parse(New StringReader(htmlText))

        '            doc.Close()
        '            rslt = ms.GetBuffer
        '        End Using

        '        If rslt.Length > 0 Then
        '            File.WriteAllBytes(TargetFile, rslt)
        '            MsgBox("Done")
        '        End If

        '    End Using
        'Catch ex As Exception
        '    Debug.Print(ex.Message)
        'End Try
#End Region


#Region " Example 4 "
        Try
            Dim pdf As Byte() = {}

            Dim cssF As String = File.ReadAllText(Path.Combine(Application.StartupPath, "table2.css"))
            Dim htmF As String = File.ReadAllText(Path.Combine(Application.StartupPath, "table2.htm"))

            Using ms As MemoryStream = New MemoryStream

                Dim doc As Document = New Document(PageSize.A4, 50, 50, 60, 60)
                Dim writer = PdfWriter.GetInstance(doc, ms)
                doc.Open()

                Using cssMemoryStream As MemoryStream = New MemoryStream(System.Text.Encoding.UTF8.GetBytes(cssF))
                    Using htmlMemoryStream As MemoryStream = New MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmF))
                        XMLWorkerHelper.GetInstance().ParseXHtml(writer, doc, htmlMemoryStream, cssMemoryStream)
                    End Using
                End Using

                doc.Close()

                pdf = ms.ToArray

            End Using

            If pdf.Length > 0 Then
                File.WriteAllBytes(TargetFile, pdf)
                MsgBox("Done")
            End If

        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try

#End Region

    End Sub

    'https://www.aspsnippets.com/Articles/Export-HTML-string-to-PDF-file-using-iTextSharp-in-ASPNet.aspx
    'Protected Sub ExportToPDF(sender As Object, e As EventArgs)
    '    Dim sr As New StringReader(Request.Form(hfGridHtml.UniqueID))
    '    Dim pdfDoc As New Document(PageSize.A4, 10.0F, 10.0F, 10.0F, 0.0F)
    '    Dim writer As PdfWriter = PdfWriter.GetInstance(pdfDoc, Response.OutputStream)
    '    pdfDoc.Open()
    '    XMLWorkerHelper.GetInstance().ParseXHtml(writer, pdfDoc, sr)
    '    pdfDoc.Close()
    '    Response.ContentType = "application/pdf"
    '    Response.AddHeader("content-disposition", "attachment;filename=HTML.pdf")
    '    Response.Cache.SetCacheability(HttpCacheability.NoCache)
    '    Response.Write(pdfDoc)
    '    Response.End()
    'End Sub



    'After installing the iTextSharp dll we can make use of the following code to generate the PDF.
    'Using (System.IO.StreamReader Reader = New System.IO.StreamReader(@"D:\test.html"))
    '   {
    '    String fileContent = Reader.ReadToEnd();
    '    Using (FileStream stream = New FileStream(@"D:\test.pdf", FileMode.Create))
    '    {
    '        Document pdfDoc = New Document(PageSize.A4);
    '        PdfWriter writer = PdfWriter.GetInstance(pdfDoc, Stream);
    '        pdfDoc.Open();
    '        StringReader sr = New StringReader(fileContent.ToString());
    '        XMLWorkerHelper.GetInstance().ParseXHtml(writer, pdfDoc, sr);
    '        pdfDoc.Close();
    '        stream.Close();
    '    }

    '}
End Class
