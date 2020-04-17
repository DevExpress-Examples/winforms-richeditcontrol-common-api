Imports Microsoft.VisualBasic
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports System
Imports DevExpress.XtraRichEdit.Services
Imports DevExpress.XtraRichEdit.Commands
Imports DevExpress.XtraBars
Imports System.Windows.Forms
Imports System.Linq
Imports DevExpress.Portable


Namespace RichEditAPISample.CodeExamples
	Public NotInheritable Class RichEditControlActions

		Private Sub New()
		End Sub
		Private Shared Sub ShowSearchFormMethod(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)
'			#Region "#ShowSearchFormMethod"
			AddHandler buttonCustomAction.ItemClick, AddressOf buttonCustomAction_ItemClick_ShowSearchFormMethod
			buttonCustomAction.Tag = richEditControl
            richEditControl.Text = "Click the 'Custom Action' button located on the RibbonControl to invoke the 'Find and Replace' dialog switched to the 'Search' tab and with the 'test' word used for search"
'			#End Region ' #ShowSearchFormMethod
		End Sub

'			#Region "#ShowSearchFormMethodCustomActionHandler"
		Private Shared Sub buttonCustomAction_ItemClick_ShowSearchFormMethod(ByVal sender As Object, ByVal e As ItemClickEventArgs)
			Dim richEdit As RichEditControl = TryCast(e.Item.Tag, RichEditControl)
			richEdit.ShowSearchForm()
            Dim searchForm As DevExpress.XtraRichEdit.Forms.SearchTextForm =
                TryCast(Application.OpenForms.Cast(Of Form)().Last(), DevExpress.XtraRichEdit.Forms.SearchTextForm)
			If searchForm IsNot Nothing Then
				Dim searchBoxes() As System.Windows.Forms.Control = searchForm.Controls.Find("cbFndSearchString", True)
                If searchBoxes.Length > 0 Then
                    searchBoxes(0).Text = "test"
                End If
			End If
		End Sub
'			#End Region ' #ShowSearchFormMethodCustomActionHandler


		Private Shared Sub ShowReplaceFormMethod(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)
'			#Region "#ShowReplaceFormMethod"
			AddHandler buttonCustomAction.ItemClick, AddressOf buttonCustomAction_ItemClick_ShowReplaceFormMethod
			buttonCustomAction.Tag = richEditControl
            richEditControl.Text = "Click the 'Custom Action' button located on the RibbonControl to invoke the 'Find and Replace' dialog switched to the 'Replace' tab"
'			#End Region ' #ShowReplaceFormMethod
		End Sub

'			#Region "#ShowReplaceFormMethodCustomActionHandler"
		Private Shared Sub buttonCustomAction_ItemClick_ShowReplaceFormMethod(ByVal sender As Object, ByVal e As ItemClickEventArgs)
			Dim richEdit As RichEditControl = TryCast(e.Item.Tag, RichEditControl)
			Dim buttonWordRanges() As DocumentRange = richEdit.Document.FindAll("button", SearchOptions.None)
			richEdit.Document.Selection = buttonWordRanges(0)
			richEdit.ShowReplaceForm()
		End Sub
'			#End Region ' #ShowReplaceFormMethodCustomActionHandler

		Private Shared Sub ShowPrintPreviewMethod(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)
'			#Region "#ShowPrintPreviewMethod"
			AddHandler buttonCustomAction.ItemClick, AddressOf buttonCustomAction_ItemClick_ShowPrintPreviewMethod
			buttonCustomAction.Tag = richEditControl
            richEditControl.Text = "Click the 'Custom Action' button located on the RibbonControl to invoke the 'Preview' window"
'			#End Region ' #ShowPrintPreviewMethod
		End Sub

'			#Region "#ShowPrintPreviewMethodCustomActionHandler"
		Private Shared Sub buttonCustomAction_ItemClick_ShowPrintPreviewMethod(ByVal sender As Object, ByVal e As ItemClickEventArgs)
			Dim richEdit As RichEditControl = TryCast(e.Item.Tag, RichEditControl)
			If richEdit.IsPrintingAvailable Then
				richEdit.ShowPrintPreview()
			End If
		End Sub
'			#End Region ' #ShowPrintPreviewMethodCustomActionHandler

		Private Shared Sub ShowPrintDialogMethod(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)
'			#Region "#ShowPrintDialogMethod"
			AddHandler buttonCustomAction.ItemClick, AddressOf buttonCustomAction_ItemClick_ShowPrintDialogMethod
			buttonCustomAction.Tag = richEditControl
			richEditControl.Text = "Click the 'Custom Action' button located on the RibbonControl to invoke the 'Print' window"
'			#End Region ' #ShowPrintDialogMethod
		End Sub

'			#Region "#ShowPrintDialogMethodCustomActionHandler"
		Private Shared Sub buttonCustomAction_ItemClick_ShowPrintDialogMethod(ByVal sender As Object, ByVal e As ItemClickEventArgs)
			Dim richEdit As RichEditControl = TryCast(e.Item.Tag, RichEditControl)
			If richEdit.IsPrintingAvailable Then
				richEdit.ShowPrintDialog()
			End If
		End Sub
'			#End Region ' #ShowPrintDialogMethodCustomActionHandler

		Private Shared Sub ScrollToCaretMethod(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)
'			#Region "#ScrollToCaretMethod"
			AddHandler buttonCustomAction.ItemClick, AddressOf buttonCustomAction_ItemClick_ScrollToCaretMethod
			buttonCustomAction.Tag = richEditControl
			richEditControl.Text = "Click the 'Custom Action' button located on the RibbonControl to load a multi-page document and scroll to the end of this document"
'			#End Region ' #ScrollToCaretMethod
		End Sub

'			#Region "#ScrollToCaretMethodCustomActionHandler"
		Private Shared Sub buttonCustomAction_ItemClick_ScrollToCaretMethod(ByVal sender As Object, ByVal e As ItemClickEventArgs)
			Dim richEdit As RichEditControl = TryCast(e.Item.Tag, RichEditControl)
			richEdit.LoadDocument("Documents\MultiPageDocument.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
			richEdit.Document.CaretPosition = richEdit.Document.Range.End
			richEdit.ScrollToCaret()
		End Sub
'			#End Region ' #ScrollToCaretMethodCustomActionHandler


		Private Shared Sub SaveDocumentMethod(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)
'			#Region "#SaveDocumentMethod"
			AddHandler buttonCustomAction.ItemClick, AddressOf buttonCustomAction_ItemClick_SaveDocumentMethod
			buttonCustomAction.Tag = richEditControl
			richEditControl.Text = "Click the 'Custom Action' button located on the RibbonControl to display a confirmation window to save a document." & Constants.vbCrLf
			richEditControl.Text &= "Click the 'Yes' button of this window to save the document to the default ('savedResults.docx') location." & Constants.vbCrLf
			richEditControl.Text &= "Click the 'No' button of this window to specify a location to save the document." & Constants.vbCrLf
'			#End Region ' #SaveDocumentMethod
		End Sub

'			#Region "#SaveDocumentMethodCustomActionHandler"
		Private Shared Sub buttonCustomAction_ItemClick_SaveDocumentMethod(ByVal sender As Object, ByVal e As ItemClickEventArgs)
			Dim richEdit As RichEditControl = TryCast(e.Item.Tag, RichEditControl)
            If MessageBox.Show("Do you want to save this document to the default ('savedResults.docx') location?",
                               "Saving a document", MessageBoxButtons.YesNo) = DialogResult.Yes Then

                richEdit.SaveDocument("savedResults.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            Else
                richEdit.SaveDocumentAs()
            End If
			System.Windows.Forms.MessageBox.Show("A document was saved sucsessfully")
		End Sub
'			#End Region ' #SaveDocumentMethodCustomActionHandler


		Private Shared Sub PrintDocumentMethod(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)
'			#Region "#PrintDocumentMethod"
			AddHandler buttonCustomAction.ItemClick, AddressOf buttonCustomAction_ItemClick_PrintDocumentMethod
			buttonCustomAction.Tag = richEditControl
			richEditControl.Text = "Click the 'Custom Action' button located on the RibbonControl to print a current document to the default printer" & Constants.vbCrLf & Constants.vbCrLf & Constants.vbCrLf
'			#End Region ' #PrintDocumentMethod
		End Sub

'			#Region "#PrintDocumentMethodCustomActionHandler"
		Private Shared Sub buttonCustomAction_ItemClick_PrintDocumentMethod(ByVal sender As Object, ByVal e As ItemClickEventArgs)
			Dim richEdit As RichEditControl = TryCast(e.Item.Tag, RichEditControl)
			If richEdit.IsPrintingAvailable Then
				richEdit.Print()
			End If
		End Sub
'			#End Region ' #PrintDocumentMethodCustomActionHandler


		Private Shared Sub PasteTextFromClipboardMethod(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)
'			#Region "#PasteTextFromClipboard"
			AddHandler buttonCustomAction.ItemClick, AddressOf buttonCustomAction_ItemClick_PasteTextFromClipboard
			buttonCustomAction.Tag = richEditControl
			richEditControl.Text = "Click the 'Custom Action' button located on the RibbonControl to paste text from the system clipboard into a current document's position" & Constants.vbCrLf & Constants.vbCrLf & Constants.vbCrLf
'			#End Region ' #PasteTextFromClipboard
		End Sub

'			#Region "#PasteTextFromClipboardCustomActionHandler"
		Private Shared Sub buttonCustomAction_ItemClick_PasteTextFromClipboard(ByVal sender As Object, ByVal e As ItemClickEventArgs)
			Dim richEdit As RichEditControl = TryCast(e.Item.Tag, RichEditControl)
			richEdit.Paste()
		End Sub
'			#End Region ' #PasteTextFromClipboardCustomActionHandler


		Private Shared Sub MailMergeMethod(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)
'			#Region "#MailMergeMethod"
			AddHandler buttonCustomAction.ItemClick, AddressOf buttonCustomAction_ItemClick_MailMergeMethod
			buttonCustomAction.Tag = richEditControl
			richEditControl.Text = "Click the 'Custom Action' button located on the RibbonControl to perform the 'Mail Merge' operation" & Constants.vbCrLf
'			#End Region ' #MailMergeMethod
		End Sub

'			#Region "#MailMergeMethodCustomActionHandler"
		Private Shared Sub buttonCustomAction_ItemClick_MailMergeMethod(ByVal sender As Object, ByVal e As ItemClickEventArgs)
			Dim richEdit As RichEditControl = TryCast(e.Item.Tag, RichEditControl)
			richEdit.LoadDocument("Documents\MailMergeSimple.rtf", DevExpress.XtraRichEdit.DocumentFormat.Rtf)
			Dim MailMergeDataSource As New System.Data.DataTable()
			MailMergeDataSource.Columns.Add("FirstName")
			MailMergeDataSource.Columns.Add("LastName")
			MailMergeDataSource.Columns.Add("HiringDate")
			MailMergeDataSource.Columns.Add("Address")
			MailMergeDataSource.Columns.Add("City")
			MailMergeDataSource.Columns.Add("Country")
			MailMergeDataSource.Columns.Add("Position")
			MailMergeDataSource.Columns.Add("CompanyName")
			MailMergeDataSource.Columns.Add("Gender")
			MailMergeDataSource.Columns.Add("Phone")
			MailMergeDataSource.Columns.Add("HRManagerName")

            Dim firstName() As String = {"Nancy", "Andrew", "Janet", "Margaret",
                                         "Steven", "Michael", "Robert", "Laura", "Anne"}

            Dim lastName() As String = {"Davolio", "Fuller", "Leverling", "Peacock",
                                        "Buchanan", "Suyama", "King", "Callahan", "Dodsworth"}

            Dim city() As String = {"Seattle", "Tacoma", "Kirkland", "Redmond",
                                    "London", "London", "London", "Seattle", "London"}

            Dim country() As String = {"USA", "USA", "USA", "USA",
                                       "UK", "UK", "UK", "USA", "UK"}

            Dim address() As String = {"507 - 20th Ave. E. Apt. 2A", "908 W. Capital Way", "722 Moss Bay Blvd.",
                                       "4110 Old Redmond Rd.", "14 Garrett Hill", "Coventry House Miner Rd.",
                                       "Edgeham Hollow Winchester Way", "4726 - 11th Ave. N.E.", "7 Houndstooth Rd."}

            Dim position() As String = {"Sales Representative", "Vice President, Sales", "Sales Representative",
                                        "Sales Representative", "Sales Manager", "Sales Representative",
                                        "Sales Representative", "Inside Sales Coordinator", "Sales Representative"}

            Dim gender() As Char = {"F"c, "M"c, "F"c, "F"c, "M"c, "M"c, "M"c, "F"c, "F"c}

            Dim phone() As String = {"(206) 555-9857", "(206) 555-9482", "(206) 555-3412", "(206) 555-8122",
                                     "(71) 555-4848", "(71) 555-7773", "(71) 555-5598", "(206) 555-1189", "(71) 555-4444"}

            Dim companyName() As String = {"Consolidated Holdings", "Around the Horn", "North/South", "Island Trading",
                                           "White Clover Markets", "Trail's Head Gourmet Provisioners", "The Cracker Box",
                                           "The Big Cheese", "Rattlesnake Canyon Grocery", "Split Rail Beer & Ale",
                                           "Hungry Coyote Import Store", "Great Lakes Food Market"}

			Dim rnd As New Random()
			For i As Integer = 0 To 8
                MailMergeDataSource.Rows.Add(New Object() {
                                             firstName(i),
                                             lastName(i),
                                             DateTime.Now.AddDays(-(rnd.Next(0, 2000))),
                                             address(i),
                                             city(i),
                                             country(i),
                                             position(i),
                                             companyName(i),
                                             gender(i),
                                             phone(i),
                                             "Dan Marino"})
			Next i

			Dim options As MailMergeOptions = richEdit.CreateMailMergeOptions()
			options.DataSource = MailMergeDataSource
			options.FirstRecordIndex = 2
			options.LastRecordIndex = 5
			options.MergeMode = MergeMode.NewSection

			Using fsResult As New System.IO.FileStream("MailMergeResult.rtf", System.IO.FileMode.OpenOrCreate)
				richEdit.MailMerge(options, fsResult, DevExpress.XtraRichEdit.DocumentFormat.Rtf)
			End Using
			System.Diagnostics.Process.Start("MailMergeResult.rtf")
		End Sub
'			#End Region ' #MailMergeMethodCustomActionHandler


		Private Shared Sub LoadDocumentTemplateMethod(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)
'			#Region "#LoadDocumentTemplate"
			AddHandler buttonCustomAction.ItemClick, AddressOf buttonCustomAction_ItemClick_LoadDocumentTemplate
			buttonCustomAction.Tag = richEditControl
			richEditControl.Text = "Click the 'Custom Action' button located on the RibbonControl to load the 'LoremDocumentTest' document as a template document." & Constants.vbCrLf
			richEditControl.Text &= "It means that if you change any text in this document and click the 'Save' button, a current document will not be overwritten "
			richEditControl.Text &= "and the 'Save file as' dialog will be shown to save the RichEditControl's document as a new document"
'			#End Region ' #LoadDocumentTemplate
		End Sub

'			#Region "#LoadDocumentTemplateCustomActionHandler"
		Private Shared Sub buttonCustomAction_ItemClick_LoadDocumentTemplate(ByVal sender As Object, ByVal e As ItemClickEventArgs)
			Dim richEdit As RichEditControl = TryCast(e.Item.Tag, RichEditControl)
			richEdit.LoadDocumentTemplate("Documents\LoremDocumentTest.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
		End Sub
'			#End Region ' #LoadDocumentTemplateCustomActionHandler


		Private Shared Sub LoadDocumentMethod(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)
'			#Region "#LoadDocumentMethod"
			AddHandler buttonCustomAction.ItemClick, AddressOf buttonCustomAction_ItemClick_LoadDocument
			buttonCustomAction.Tag = richEditControl
			richEditControl.Text = "Click the 'Custom Action' button located on the RibbonControl " & Constants.vbCrLf
			richEditControl.Text &= "to sequentially load documents in different formats (RTF, DOCX, HTML) using the RichEditDocumentServer " & Constants.vbCrLf
			richEditControl.Text &= "and append the loaded document contents to a current RichEditControl's document"
'			#End Region ' #LoadDocumentMethod
		End Sub

'			#Region "#LoadDocumentMethodCustomActionHandler"
		Private Shared Sub buttonCustomAction_ItemClick_LoadDocument(ByVal sender As Object, ByVal e As ItemClickEventArgs)
			Dim richEdit As RichEditControl = TryCast(e.Item.Tag, RichEditControl)

			Dim documentServer As IRichEditDocumentServer = richEdit.CreateDocumentServer()
			Using fs As New System.IO.FileStream("Documents\testDocumentDOCX.docx", System.IO.FileMode.Open)
				documentServer.LoadDocument(fs, DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
				richEdit.Document.AppendDocumentContent(documentServer.Document.Range)
			End Using

			Using fs As New System.IO.FileStream("Documents\testDocumentRTF.rtf", System.IO.FileMode.Open)
				documentServer.LoadDocument(fs, DevExpress.XtraRichEdit.DocumentFormat.Rtf)
				richEdit.Document.AppendDocumentContent(documentServer.Document.Range)
			End Using

			Using fs As New System.IO.FileStream("Documents\tesDocumentHTML.html", System.IO.FileMode.Open)
				documentServer.LoadDocument(fs, DevExpress.XtraRichEdit.DocumentFormat.Html)
				richEdit.Document.AppendDocumentContent(documentServer.Document.Range)
			End Using
		End Sub
'			#End Region ' #LoadDocumentMethodCustomActionHandler


		Private Shared Sub IsSelectionInTable(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)
'			#Region "#IsSelectionInTable"
			richEditControl.Tag = buttonCustomAction
			AddHandler richEditControl.SelectionChanged, AddressOf richEditControl_SelectionChanged
			richEditControl.LoadDocument("Documents\TableSampleDocument.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
'			#End Region ' #IsSelectionInTable
		End Sub

'			#Region "#IsSelectionInTableAdditionalModules"
		Private Shared Sub richEditControl_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs)
			Dim richEdit As RichEditControl = (TryCast(sender, RichEditControl))
			Dim barButton As BarButtonItem = TryCast(richEdit.Tag, BarButtonItem)
			barButton.Enabled = Not richEdit.IsSelectionInTable()
		End Sub
'			#End Region ' #IsSelectionInTableAdditionalModules


		Private Shared Sub GetPositionFromPoint(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)
'			#Region "#GetPositionFromPoint"
			AddHandler richEditControl.MouseMove, AddressOf richEditControl1_MouseMove
			richEditControl.LoadDocument("Documents\LoremDocumentTest.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
			richEditControl.Document.InsertText(richEditControl.Document.Range.Start, "Hovering a mouse in this document results in showing a tooltip with a current document position and a character in this position" & Constants.vbCrLf & Constants.vbCrLf)
            '			#End Region ' #GetPositionFromPoint
		End Sub


'			#Region "#GetPositionFromPointAdditionalModules"
		Private Shared testToolTipController As New DevExpress.Utils.ToolTipController()

		Private Shared Sub richEditControl1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
			Dim richEdit As RichEditControl = TryCast(sender, RichEditControl)
            Dim docPoint As System.Drawing.Point =
                DevExpress.Office.Utils.Units.PixelsToDocuments(e.Location, richEdit.DpiX, richEdit.DpiY)

			Dim pos As DocumentPosition = richEdit.GetPositionFromPoint(docPoint)
			If pos Is Nothing Then
				Return
			End If
            Dim currentToolTipText As String = String.Format("Position: {0}, Character: {1}", pos.ToString(),
                                                             richEdit.Document.GetText(richEdit.Document.CreateRange(pos, 1)))

			Dim info As New DevExpress.Utils.ToolTipControlInfo(currentToolTipText, currentToolTipText)
			info.ToolTipPosition = System.Windows.Forms.Form.MousePosition
			testToolTipController.ShowHint(info)
		End Sub
'			#End Region ' #GetPositionFromPointAdditionalModules

		Private Shared Sub GetBoundsFromDocumentPosition(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)
'			#Region "#GetBoundsFromDocumentPosition"
			AddHandler richEditControl.Paint, AddressOf richEditControl_Paint
			richEditControl.LoadDocument("Documents\LoremDocumentTest.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            richEditControl.Document.InsertText(richEditControl.Document.Range.Start, "A document row which contains a caret position is highlighted" & Constants.vbCrLf & Constants.vbCrLf)
'			#End Region ' #GetBoundsFromDocumentPosition
		End Sub

'			#Region "#GetBoundsFromDocumentPositionAdditionalModules"
		Private Shared Sub richEditControl_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs)
			Dim richEdit As RichEditControl = TryCast(sender, RichEditControl)
			Dim pos As DocumentPosition = richEdit.Document.CaretPosition
			If pos IsNot Nothing Then
                Dim rect As System.Drawing.Rectangle = DevExpress.Office.Utils.Units.DocumentsToPixels(
                    richEdit.GetBoundsFromPosition(pos),
                    richEdit.DpiX,
                    richEdit.DpiY)

				Dim firstSection As Section = richEdit.Document.Sections(0)
				Dim pageWidth As Integer = Convert.ToInt32(firstSection.Page.Width - firstSection.Margins.Left - firstSection.Margins.Right)
                e.Graphics.DrawLine(System.Drawing.Pens.Red,
                                    New System.Drawing.Point(0, rect.Bottom),
                                    New System.Drawing.Point(pageWidth, rect.Bottom))
			End If
		End Sub
'			#End Region ' #GetBoundsFromDocumentPositionAdditionalModules

	   Shared Sub ExportDocumentToPdfFormat(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)
'			#Region "#ExportDocumentToPdfFormat"
			AddHandler buttonCustomAction.ItemClick, AddressOf buttonCustomAction_ItemClick_PDF
			buttonCustomAction.Tag = richEditControl
			richEditControl.Text = "Click the 'Custom Action' button located in the ribbon "
			richEditControl.Text &= "to load the 'Grimm.docx' file into RichEditControl's document and export the document into PDF format" & ControlChars.CrLf
'			#End Region ' #ExportDocumentToPdfFormat
	   End Sub

		#Region "#ExportDocumentToPdfFormatCustomActionHandler"
		Shared Sub buttonCustomAction_ItemClick_PDF(ByVal sender As Object, ByVal e As ItemClickEventArgs)
			Dim richEdit As RichEditControl = TryCast(e.Item.Tag, RichEditControl)
			richEdit.LoadDocument("Documents\Grimm.docx")
			'Set the required export options:
			Dim options As New DevExpress.XtraPrinting.PdfExportOptions()
			options.DocumentOptions.Author = "Mark Jones"
			options.Compressed = False
			options.ImageQuality = DevExpress.XtraPrinting.PdfJpegImageQuality.High
			'Export the document to the file:
			richEdit.ExportToPdf("resultingDocument.pdf", options)
			'Export the document to the file stream:
			Using pdfFileStream As New IO.FileStream("resultingDocumentFromStream.pdf", IO.FileMode.Create)
				richEdit.ExportToPdf(pdfFileStream, options)
			End Using

			System.Diagnostics.Process.Start("resultingDocument.pdf")
		End Sub
		#End Region ' #ExportDocumentToPdfFormatCustomActionHandler

		Shared Sub ExportDocumentToHtmlFormat(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)
'			#Region "#ExportDocumentToHtmlFormat"
			AddHandler buttonCustomAction.ItemClick, AddressOf buttonCustomAction_ItemClick_Html
			buttonCustomAction.Tag = richEditControl
			richEditControl.Text = "Click the 'Custom Action' button located in the ribbon "
			richEditControl.Text &= "to load the 'Grimm.docx' file into RichEditControl's document and export the document into HTML format" & ControlChars.CrLf
'			#End Region ' #ExportDocumentToHtmlFormat
		End Sub

		#Region "#ExportDocumentToHtmlFormatCustomActionHandler"
		Shared Sub buttonCustomAction_ItemClick_Html(ByVal sender As Object, ByVal e As ItemClickEventArgs)
			Dim richEdit As RichEditControl = TryCast(e.Item.Tag, RichEditControl)
			richEdit.LoadDocument("Documents\Grimm.docx")
			'Export document to the file:
			richEdit.SaveDocument("resultingDocument.html", DocumentFormat.Html)
			'Export document to the stream:
			Using htmlFileStream As New IO.FileStream("Document_HTML.html", IO.FileMode.Create)
				richEdit.SaveDocument(htmlFileStream, DocumentFormat.Html)
			End Using

			System.Diagnostics.Process.Start("resultingDocument.html")
		End Sub
		#End Region ' #ExportDocumentToHtmlFormatCustomActionHandler


		Private Shared Sub DeselectAllTextInDocument(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)
'			#Region "#DeselectAllTextInDocument"
			AddHandler buttonCustomAction.ItemClick, AddressOf buttonCustomAction_ItemClick_Deselect
			buttonCustomAction.Tag = richEditControl
			richEditControl.Text = "Select text in a document and сlick the 'Custom Action' button located on the RibbonControl." & Constants.vbCrLf
			richEditControl.Text &= "As a result, the previously selected text is deselcted and the caret position is set to the end of the text" & Constants.vbCrLf
'			#End Region ' #DeselectAllTextInDocument
		End Sub

'			#Region "#DeselectAllTextInDocumentCustomActionHandler"
		Private Shared Sub buttonCustomAction_ItemClick_Deselect(ByVal sender As Object, ByVal e As ItemClickEventArgs)
			Dim richEdit As RichEditControl = TryCast(e.Item.Tag, RichEditControl)
			Dim endOfSelection As DocumentPosition = richEdit.Document.Selection.End
			richEdit.DeselectAll()
			richEdit.Document.CaretPosition = endOfSelection
		End Sub
'			#End Region ' #DeselectAllTextInDocumentCustomActionHandler

		Private Shared Sub CutSelectedText(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)
'			#Region "#CutSelectedText"
			AddHandler buttonCustomAction.ItemClick, AddressOf buttonCustomAction_ItemClick_Cut
			buttonCustomAction.Tag = richEditControl
			richEditControl.Text = "Click the 'Custom Action' button located on the RibbonControl "
			richEditControl.Text &= "to cut the 'Custom Action' text block " & Constants.vbCrLf
'			#End Region ' #CutSelectedText
		End Sub

'			#Region "#CutSelectedTextCustomActionHandler"
		Private Shared Sub buttonCustomAction_ItemClick_Cut(ByVal sender As Object, ByVal e As ItemClickEventArgs)
			Dim richEdit As RichEditControl = TryCast(e.Item.Tag, RichEditControl)
			Dim foundedRanges() As DocumentRange = richEdit.Document.FindAll("Custom Action", SearchOptions.None)
			For i As Integer = 0 To foundedRanges.Length - 1
				richEdit.Document.Selection = foundedRanges(i)
				richEdit.Cut()
			Next i
		End Sub
'			#End Region ' #CutSelectedTextCustomActionHandler


        Private Shared Sub CreateNewDocument(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)
            '			#Region "#CreateNewDocument"
            AddHandler buttonCustomAction.ItemClick, AddressOf buttonCustomAction_ItemClick_NewDocument
            buttonCustomAction.Tag = richEditControl
            richEditControl.Text = "Click the 'Custom Action' button located on the RibbonControl "
            richEditControl.Text &= "to create a new document " & Constants.vbCrLf
            '			#End Region ' #CreateNewDocument
        End Sub

        Private Shared Sub ReplaceSelectedText(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)
            '			#Region "#ReplaceSelectedText"
            AddHandler buttonCustomAction.ItemClick, AddressOf buttonCustomAction_ItemClick_Replace
            buttonCustomAction.Tag = richEditControl
            richEditControl.Text = "Click the 'Custom Action' button located on the RibbonControl "
            richEditControl.Text &= "to replace selected text with asterisks." & ControlChars.CrLf
            '			#End Region ' #ReplaceSelectedText
        End Sub

        '			#Region "#ReplaceSelectedTextCustomActionHandler"
        Private Shared Sub buttonCustomAction_ItemClick_Replace(ByVal sender As Object, ByVal e As ItemClickEventArgs)
            Dim richEdit As RichEditControl = TryCast(e.Item.Tag, RichEditControl)
            Dim range As DocumentRange = richEdit.Document.Selection
            Dim selLength As Integer = range.Length
            Dim s As New String("*"c, selLength)
            Dim doc As SubDocument = range.BeginUpdateDocument()
            doc.InsertSingleLineText(range.Start, s)
            Dim rangeToRemove As DocumentRange = doc.CreateRange(range.Start, selLength)
            doc.Delete(rangeToRemove)
            range.EndUpdateDocument(doc)
        End Sub
        '			#End Region '#ReplaceSelectedTextCustomActionHandler


        '			#Region "#CreateNewDocumentCustomActionHandler"
        Private Shared Sub buttonCustomAction_ItemClick_NewDocument(ByVal sender As Object, ByVal e As ItemClickEventArgs)
			Dim richEdit As RichEditControl = TryCast(e.Item.Tag, RichEditControl)
			richEdit.CreateNewDocument()
		End Sub
        '			#End Region '#CreateNewDocumentCustomActionHandler

        Private Shared Sub CreateDocumentServer(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)
'			#Region "#CreateDocumentServer"
			AddHandler buttonCustomAction.ItemClick, AddressOf buttonCustomAction_ItemClick_DocumentServer
			buttonCustomAction.Tag = richEditControl
			richEditControl.Text = "Click the 'Custom Action' button located on the RibbonControl "
			richEditControl.Text &= "to insert a content of the 'DocumentServerTest.docx' file into a current document's position " & Constants.vbCrLf
'			#End Region ' #CreateDocumentServer
		End Sub

'			#Region "#CreateDocumentServerCustomActionHandler"
		Private Shared Sub buttonCustomAction_ItemClick_DocumentServer(ByVal sender As Object, ByVal e As ItemClickEventArgs)
			Dim richEdit As RichEditControl = TryCast(e.Item.Tag, RichEditControl)
			Dim server As IRichEditDocumentServer = richEdit.CreateDocumentServer()
			Using fs As New System.IO.FileStream("Documents\DocumentServerTest.docx", System.IO.FileMode.Open)
				server.LoadDocument(fs, DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
				richEdit.Document.InsertDocumentContent(richEdit.Document.CaretPosition, server.Document.Range)
			End Using
		End Sub
'			#End Region ' #CreateDocumentServerCustomActionHandler

		Private Shared Sub CreateAndExecuteCommands(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)
'			#Region "#CreateAndExecuteCommands"
			AddHandler buttonCustomAction.ItemClick, AddressOf buttonCustomAction_ItemClick_Commands
			buttonCustomAction.Tag = richEditControl
			richEditControl.Text = "Click the 'Custom Action' button located on the RibbonControl to capitalize each word in this text," & Constants.vbCrLf
			richEditControl.Text &= "to change font of this text to Bold and show the Print Preview window with a result." & Constants.vbCrLf
			richEditControl.Text &= "All these actions are performed using corresponding RichEditControl commands." & Constants.vbCrLf
'			#End Region ' #CreateAndExecuteCommands
		End Sub

'			#Region "#CreateAndExecuteCommandsCustomActionHandler"
		Private Shared Sub buttonCustomAction_ItemClick_Commands(ByVal sender As Object, ByVal e As ItemClickEventArgs)
			Dim richEdit As RichEditControl = TryCast(e.Item.Tag, RichEditControl)
			richEdit.SelectAll()

			Dim capCommand As RichEditCommand = richEdit.CreateCommand(RichEditCommandId.CapitalizeEachWordTextCase)
			capCommand.ForceExecute(capCommand.CreateDefaultCommandUIState())

			Dim boldCommand As RichEditCommand = richEdit.CreateCommand(RichEditCommandId.ToggleFontBold)
            boldCommand.ForceExecute(boldCommand.CreateDefaultCommandUIState())

            Dim changeFontColorCommand As RichEditCommand = richEdit.CreateCommand(RichEditCommandId.ChangeFontBackColor)
            Dim state As DevExpress.Utils.Commands.ICommandUIState = changeFontColorCommand.CreateDefaultCommandUIState()
            state.EditValue = System.Drawing.Color.Yellow
            changeFontColorCommand.ForceExecute(state)


            Dim previewCommand As RichEditCommand = richEdit.CreateCommand(RichEditCommandId.PrintPreview)
            previewCommand.ForceExecute(previewCommand.CreateDefaultCommandUIState())

            richEdit.DeselectAll()
		End Sub
'			#End Region ' #CreateAndExecuteCommandsCustomActionHandler

		Private Shared Sub CopySelectedText(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)
'			#Region "#CopySelectedText"
            richEditControl.Text = "Open any text editor (Notepad, Word) and press the 'Ctrl+V' shortcut, the 'text from RichEditControl' string will be pasted into the editor's document"
			Dim foundedRanges() As DocumentRange = richEditControl.Document.FindAll("text from RichEditControl", SearchOptions.None)
			richEditControl.Document.Selection = foundedRanges(0)
			richEditControl.Copy()
'			#End Region ' #CopySelectedText
		End Sub

		Private Shared Sub CustomizePredefinedShortcutKeys(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)
'			#Region "#CustomizePredefinedShortcutKeys"
            richEditControl.Text = "Use 'Ctrl+G' shortcut to show/hide whitespace characters"
			richEditControl.Text += Constants.vbCrLf & "A new document can't be created by pressing 'Ctrl+N' shortcut since this shortcut is disabled"
			richEditControl.AssignShortcutKeyToCommand(System.Windows.Forms.Keys.Control, System.Windows.Forms.Keys.G, RichEditCommandId.ToggleShowWhitespace)
			richEditControl.RemoveShortcutKey(System.Windows.Forms.Keys.Control, System.Windows.Forms.Keys.N)
'			#End Region ' #CustomizePredefinedShortcutKeys
		End Sub

		Private Shared Sub ReplacePredefinedService(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)
			richEditControl.ActiveViewType = RichEditViewType.PrintLayout
'			#Region "#ReplacePredefinedService"
			richEditControl.Text = "A message box is displayed when a document is saved on the 'Save' or 'Save As' button click since custom commands are used"
            Dim commandFactory = New CustomRichEditCommandFactoryService(richEditControl, richEditControl.GetService(Of IRichEditCommandFactoryService)())
            richEditControl.ReplaceService(Of IRichEditCommandFactoryService)(commandFactory)
            '			#End Region ' #ReplacePredefinedService
        End Sub

'			#Region "#ReplacePredefinedServiceAdditionalModules"
		Public Class CustomRichEditCommandFactoryService
			Implements IRichEditCommandFactoryService
			Private ReadOnly service As IRichEditCommandFactoryService
			Private ReadOnly control As RichEditControl

			Public Sub New(ByVal control As RichEditControl, ByVal service As IRichEditCommandFactoryService)
				DevExpress.Utils.Guard.ArgumentNotNull(control, "control")
				DevExpress.Utils.Guard.ArgumentNotNull(service, "service")
				Me.control = control
				Me.service = service
			End Sub

            Public Function CreateCommand(ByVal id As RichEditCommandId) As RichEditCommand Implements IRichEditCommandFactoryService.CreateCommand
                If id.Equals(RichEditCommandId.FileSaveAs) Then
                    Return New CustomSaveDocumentAsCommand(control)
                End If
                If id.Equals(RichEditCommandId.FileSave) Then
                    Return New CustomSaveDocumentCommand(control)
                End If
                Return service.CreateCommand(id)
            End Function
		End Class

		Public Class CustomSaveDocumentCommand
			Inherits SaveDocumentCommand
			Public Sub New(ByVal richEdit As IRichEditControl)
				MyBase.New(richEdit)
			End Sub

			Protected Overrides Sub ExecuteCore()
                MyBase.ExecuteCore()
                If (Not TryCast(Me.Control, RichEditControl).Modified) Then
                    MessageBox.Show("Document is saved successfully")
                End If
            End Sub
		End Class
        
		Public Class CustomSaveDocumentAsCommand
			Inherits SaveDocumentAsCommand
			Public Sub New(ByVal richEdit As IRichEditControl)
				MyBase.New(richEdit)
			End Sub

			Protected Overrides Sub ExecuteCore()
                TryCast(Me.Control, RichEditControl).Modified = True
                MyBase.ExecuteCore()
                If (Not TryCast(Me.Control, RichEditControl).Modified) Then
                    MessageBox.Show("Document is saved successfully")
                End If
            End Sub
		End Class
'			#End Region ' #ReplacePredefinedServiceAdditionalModules

    		Private Shared Sub SimpleViewLineNumbering(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)
'			#Region "#SimpleViewLineNumbering"
			richEditControl.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
			richEditControl.Views.SimpleView.Padding = New PortablePadding(60, 4, 4, 0)
			richEditControl.Views.SimpleView.AllowDisplayLineNumbers = True
			richEditControl.ActiveViewType = RichEditViewType.Simple
			richEditControl.Document.Sections(0).LineNumbering.CountBy = 1
			richEditControl.Document.CharacterStyles("Line Number").ForeColor = System.Drawing.Color.LightGray
			richEditControl.Document.CharacterStyles("Line Number").Bold = True
'			#End Region ' #SimpleViewLineNumbering

		End Sub
    End Class

End Namespace
