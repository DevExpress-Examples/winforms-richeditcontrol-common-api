using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System;
using DevExpress.XtraRichEdit.Services;
using DevExpress.XtraRichEdit.Commands;
using DevExpress.XtraBars;
using System.Windows.Forms;
using System.Linq;
using System.IO;
using System.Drawing;
using DevExpress.Portable;

namespace RichEditAPISample.CodeExamples {
    public static class RichEditControlActions {

        static void ShowSearchFormMethod(RichEditControl richEditControl, BarButtonItem buttonCustomAction) {
            #region #ShowSearchFormMethod
            buttonCustomAction.ItemClick += buttonCustomAction_ItemClick_ShowSearchFormMethod;
            buttonCustomAction.Tag = richEditControl;
            richEditControl.Text = "Click the 'Custom Action' button located in the ribbon to invoke the 'Find and Replace' dialog switched to the 'Find' tab and with the 'test' word used for search";
            #endregion #ShowSearchFormMethod
        }

        #region #ShowSearchFormMethodCustomActionHandler
        static void buttonCustomAction_ItemClick_ShowSearchFormMethod(object sender, ItemClickEventArgs e) {
            RichEditControl richEdit = e.Item.Tag as RichEditControl;
            richEdit.ShowSearchForm();
            var searchForm = Application.OpenForms.Cast<Form>().Last() as DevExpress.XtraRichEdit.Forms.SearchTextForm;
            if(searchForm != null) {
                System.Windows.Forms.Control[] searchBoxes = searchForm.Controls.Find("cbFndSearchString", true);
                if(searchBoxes.Length > 0) searchBoxes[0].Text = "test";
            }
        }
        #endregion #ShowSearchFormMethodCustomActionHandler


        static void ShowReplaceFormMethod(RichEditControl richEditControl, BarButtonItem buttonCustomAction) {
            #region #ShowReplaceFormMethod
            buttonCustomAction.ItemClick += buttonCustomAction_ItemClick_ShowReplaceFormMethod;
            buttonCustomAction.Tag = richEditControl;
            richEditControl.Text = "Click the 'Custom Action' button located in the ribbon to invoke the 'Find and Replace' dialog switched to the 'Replace' tab";
            #endregion #ShowReplaceFormMethod
        }

        #region #ShowReplaceFormMethodCustomActionHandler
        static void buttonCustomAction_ItemClick_ShowReplaceFormMethod(object sender, ItemClickEventArgs e) {
            RichEditControl richEdit = e.Item.Tag as RichEditControl;
            DocumentRange[] buttonWordRanges = richEdit.Document.FindAll("button", SearchOptions.None);
            richEdit.Document.Selection = buttonWordRanges[0];
            richEdit.ShowReplaceForm();
        }
        #endregion #ShowReplaceFormMethodCustomActionHandler

        static void ShowPrintPreviewMethod(RichEditControl richEditControl, BarButtonItem buttonCustomAction) {
            #region #ShowPrintPreviewMethod
            buttonCustomAction.ItemClick += buttonCustomAction_ItemClick_ShowPrintPreviewMethod;
            buttonCustomAction.Tag = richEditControl;
            richEditControl.Text = "Click the 'Custom Action' button located in the ribbon to invoke the 'Preview' window";
            #endregion #ShowPrintPreviewMethod
        }

        #region #ShowPrintPreviewMethodCustomActionHandler
        static void buttonCustomAction_ItemClick_ShowPrintPreviewMethod(object sender, ItemClickEventArgs e) {
            RichEditControl richEdit = e.Item.Tag as RichEditControl;
            if(richEdit.IsPrintingAvailable) {
                richEdit.ShowPrintPreview();
            }
        }
        #endregion #ShowPrintPreviewMethodCustomActionHandler

        static void ShowPrintDialogMethod(RichEditControl richEditControl, BarButtonItem buttonCustomAction) {
            #region #ShowPrintDialogMethod
            buttonCustomAction.ItemClick += buttonCustomAction_ItemClick_ShowPrintDialogMethod;
            buttonCustomAction.Tag = richEditControl;
            richEditControl.Text = "Click the 'Custom Action' button located in the ribbon to invoke the 'Print' window";
            #endregion #ShowPrintDialogMethod
        }

        #region #ShowPrintDialogMethodCustomActionHandler
        static void buttonCustomAction_ItemClick_ShowPrintDialogMethod(object sender, ItemClickEventArgs e) {
            RichEditControl richEdit = e.Item.Tag as RichEditControl;
            if(richEdit.IsPrintingAvailable) {
                richEdit.ShowPrintDialog();
            }
        }
        #endregion #ShowPrintDialogMethodCustomActionHandler

        static void ScrollToCaretMethod(RichEditControl richEditControl, BarButtonItem buttonCustomAction) {
            #region #ScrollToCaretMethod
            buttonCustomAction.ItemClick += buttonCustomAction_ItemClick_ScrollToCaretMethod;
            buttonCustomAction.Tag = richEditControl;
            richEditControl.Text = "Click the 'Custom Action' button located in the ribbon to load a multi-page document and scroll to the end of this document";
            #endregion #ScrollToCaretMethod
        }

        #region #ScrollToCaretMethodCustomActionHandler
        static void buttonCustomAction_ItemClick_ScrollToCaretMethod(object sender, ItemClickEventArgs e) {
            RichEditControl richEdit = e.Item.Tag as RichEditControl;
            richEdit.LoadDocument("Documents\\MultiPageDocument.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml);
            richEdit.Document.CaretPosition = richEdit.Document.Range.End;
            richEdit.ScrollToCaret();
        }
        #endregion #ScrollToCaretMethodCustomActionHandler


        static void SaveDocumentMethod(RichEditControl richEditControl, BarButtonItem buttonCustomAction) {
            #region #SaveDocumentMethod
            buttonCustomAction.ItemClick += buttonCustomAction_ItemClick_SaveDocumentMethod;
            buttonCustomAction.Tag = richEditControl;
            richEditControl.Text = "Click the 'Custom Action' button located in the ribbon to display a confirmation window to save a document.\r\n";
            richEditControl.Text += "Click the 'Yes' button of this window to save the document to the default ('savedResults.docx') location.\r\n";
            richEditControl.Text += "Click the 'No' button of this window to specify a location to save the document.\r\n";
            #endregion #SaveDocumentMethod
        }

        #region #SaveDocumentMethodCustomActionHandler
        static void buttonCustomAction_ItemClick_SaveDocumentMethod(object sender, ItemClickEventArgs e) {
            RichEditControl richEdit = e.Item.Tag as RichEditControl;
            if(MessageBox.Show("Do you want to save this document to the default ('savedResults.docx') location?", 
                "Saving a document", MessageBoxButtons.YesNo) == DialogResult.Yes)

                richEdit.SaveDocument("savedResults.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml);
            else
                richEdit.SaveDocumentAs();
            System.Windows.Forms.MessageBox.Show("A document was saved sucsessfully");
        }
        #endregion #SaveDocumentMethodCustomActionHandler


        static void PrintDocumentMethod(RichEditControl richEditControl, BarButtonItem buttonCustomAction) {
            #region #PrintDocumentMethod
            buttonCustomAction.ItemClick += buttonCustomAction_ItemClick_PrintDocumentMethod;
            buttonCustomAction.Tag = richEditControl;
            richEditControl.Text = "Click the 'Custom Action' button located in the ribbon to print a current document to the default printer\r\n\r\n\r\n";
            #endregion #PrintDocumentMethod
        }

        #region #PrintDocumentMethodCustomActionHandler
        static void buttonCustomAction_ItemClick_PrintDocumentMethod(object sender, ItemClickEventArgs e) {
            RichEditControl richEdit = e.Item.Tag as RichEditControl;
            if(richEdit.IsPrintingAvailable) {
                richEdit.Print();
            }
        }
        #endregion #PrintDocumentMethodCustomActionHandler


        static void PasteTextFromClipboardMethod(RichEditControl richEditControl, BarButtonItem buttonCustomAction) {
            #region #PasteTextFromClipboard
            buttonCustomAction.ItemClick += buttonCustomAction_ItemClick_PasteTextFromClipboard;
            buttonCustomAction.Tag = richEditControl;
            richEditControl.Text = "Click the 'Custom Action' button located in the ribbon to paste text from the system clipboard into a current document's position\r\n\r\n\r\n";
            #endregion #PasteTextFromClipboard
        }

        #region #PasteTextFromClipboardCustomActionHandler
        static void buttonCustomAction_ItemClick_PasteTextFromClipboard(object sender, ItemClickEventArgs e) {
            RichEditControl richEdit = e.Item.Tag as RichEditControl;
            richEdit.Paste();
        }
        #endregion #PasteTextFromClipboardCustomActionHandler


        static void MailMergeMethod(RichEditControl richEditControl, BarButtonItem buttonCustomAction) {
            #region #MailMergeMethod
            buttonCustomAction.ItemClick += buttonCustomAction_ItemClick_MailMergeMethod;
            buttonCustomAction.Tag = richEditControl;
            richEditControl.Text = "Click the 'Custom Action' button located in the ribbon to perform the 'Mail Merge' operation\r\n";
            #endregion #MailMergeMethod
        }

        #region #MailMergeMethodCustomActionHandler
        static void buttonCustomAction_ItemClick_MailMergeMethod(object sender, ItemClickEventArgs e) {
            RichEditControl richEdit = e.Item.Tag as RichEditControl;
            richEdit.LoadDocument("Documents\\MailMergeSimple.rtf", DevExpress.XtraRichEdit.DocumentFormat.Rtf);
            System.Data.DataTable MailMergeDataSource = new System.Data.DataTable();
            MailMergeDataSource.Columns.Add("FirstName");
            MailMergeDataSource.Columns.Add("LastName");
            MailMergeDataSource.Columns.Add("HiringDate");
            MailMergeDataSource.Columns.Add("Address");
            MailMergeDataSource.Columns.Add("City");
            MailMergeDataSource.Columns.Add("Country");
            MailMergeDataSource.Columns.Add("Position");
            MailMergeDataSource.Columns.Add("CompanyName");
            MailMergeDataSource.Columns.Add("Gender");
            MailMergeDataSource.Columns.Add("Phone");
            MailMergeDataSource.Columns.Add("HRManagerName");

            string[] firstName = { "Nancy", "Andrew", "Janet", "Margaret", 
                                     "Steven", "Michael", "Robert", "Laura", "Anne" };
            
            string[] lastName = { "Davolio", "Fuller", "Leverling", "Peacock", 
                                    "Buchanan", "Suyama", "King", "Callahan", "Dodsworth" };
            
            string[] city = { "Seattle", "Tacoma", "Kirkland", "Redmond", "London",                                 
                                "London", "London", "Seattle", "London" };
            
            string[] country = { "USA", "USA", "USA", "USA", 
                                   "UK", "UK", "UK", "USA", "UK" };

            string[] address = { "507 - 20th Ave. E. Apt. 2A", "908 W. Capital Way", "722 Moss Bay Blvd.", 
                                   "4110 Old Redmond Rd.", "14 Garrett Hill", "Coventry House Miner Rd.", 
                                   "Edgeham Hollow Winchester Way", "4726 - 11th Ave. N.E.", "7 Houndstooth Rd." };

            string[] position = { "Sales Representative", "Vice President, Sales", "Sales Representative", 
                                    "Sales Representative", "Sales Manager", "Sales Representative", 
                                    "Sales Representative", "Inside Sales Coordinator", "Sales Representative" };
            
            char[] gender = { 'F', 'M', 'F', 'F', 'M', 'M', 'M', 'F', 'F' };
            string[] phone = { "(206) 555-9857", "(206) 555-9482", "(206) 555-3412", "(206) 555-8122", 
                                 "(71) 555-4848", "(71) 555-7773", "(71) 555-5598", "(206) 555-1189", "(71) 555-4444" };

            string[] companyName = { "Consolidated Holdings", "Around the Horn", "North/South", "Island Trading", 
                                       "White Clover Markets", "Trail's Head Gourmet Provisioners", "The Cracker Box", 
                                       "The Big Cheese", "Rattlesnake Canyon Grocery", "Split Rail Beer & Ale", 
                                       "Hungry Coyote Import Store", "Great Lakes Food Market" };

            Random rnd = new Random();
            for(int i = 0; i < 9; i++) {
                MailMergeDataSource.Rows.Add(new object[] { 
                    firstName[i], 
                    lastName[i], 
                    DateTime.Now.AddDays(-(rnd.Next(0, 2000))), 
                    address[i], 
                    city[i], 
                    country[i], 
                    position[i], 
                    companyName[i], 
                    gender[i], 
                    phone[i], 
                    "Dan Marino" });
            }

            MailMergeOptions options = richEdit.CreateMailMergeOptions();
            options.DataSource = MailMergeDataSource;
            options.FirstRecordIndex = 2;
            options.LastRecordIndex = 5;
            options.MergeMode = MergeMode.NewSection;

            using(FileStream fsResult = new FileStream("MailMergeResult.rtf", FileMode.OpenOrCreate)) {
                richEdit.MailMerge(options, fsResult, DevExpress.XtraRichEdit.DocumentFormat.Rtf);
            }
            System.Diagnostics.Process.Start("MailMergeResult.rtf");
        }
        #endregion #MailMergeMethodCustomActionHandler


        static void LoadDocumentTemplateMethod(RichEditControl richEditControl, BarButtonItem buttonCustomAction) {
            #region #LoadDocumentTemplate
            buttonCustomAction.ItemClick += buttonCustomAction_ItemClick_LoadDocumentTemplate;
            buttonCustomAction.Tag = richEditControl;
            richEditControl.Text = "Click the 'Custom Action' button located in the ribbon to load the 'LoremDocumentTest' document as a template document.\r\n";
            richEditControl.Text += "It means that if you change any text in this document and click the 'Save' button, a current document will not be overwritten ";
            richEditControl.Text += "and the 'Save file as' dialog will be shown to save the RichEditControl's document as a new document";
            #endregion #LoadDocumentTemplate
        }

        #region #LoadDocumentTemplateCustomActionHandler
        static void buttonCustomAction_ItemClick_LoadDocumentTemplate(object sender, ItemClickEventArgs e) {
            RichEditControl richEdit = e.Item.Tag as RichEditControl;
            richEdit.LoadDocumentTemplate("Documents\\LoremDocumentTest.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml);
        }
        #endregion #LoadDocumentTemplateCustomActionHandler


        static void LoadDocumentMethod(RichEditControl richEditControl, BarButtonItem buttonCustomAction) {
            #region #LoadDocumentMethod
            buttonCustomAction.ItemClick += buttonCustomAction_ItemClick_LoadDocument;
            buttonCustomAction.Tag = richEditControl;
            richEditControl.Text = "Click the 'Custom Action' button located in the ribbon \r\n";
            richEditControl.Text += "to sequentially load documents in different formats (RTF, DOCX, HTML) using the RichEditDocumentServer \r\n";
            richEditControl.Text += "and append the loaded document contents to a current RichEditControl's document";
            #endregion #LoadDocumentMethod
        }

        #region #LoadDocumentMethodCustomActionHandler
        static void buttonCustomAction_ItemClick_LoadDocument(object sender, ItemClickEventArgs e) {
            RichEditControl richEdit = e.Item.Tag as RichEditControl;

            IRichEditDocumentServer documentServer = richEdit.CreateDocumentServer();
            using(FileStream fs = new FileStream("Documents\\testDocumentDOCX.docx", FileMode.Open)) {
                documentServer.LoadDocument(fs, DevExpress.XtraRichEdit.DocumentFormat.OpenXml);
                richEdit.Document.AppendDocumentContent(documentServer.Document.Range);
            }

            using(FileStream fs = new FileStream("Documents\\testDocumentRTF.rtf", FileMode.Open)) {
                documentServer.LoadDocument(fs, DevExpress.XtraRichEdit.DocumentFormat.Rtf);
                richEdit.Document.AppendDocumentContent(documentServer.Document.Range);
            }

            using(FileStream fs = new FileStream("Documents\\testDocumentHTML.html", FileMode.Open)) {
                documentServer.LoadDocument(fs, DevExpress.XtraRichEdit.DocumentFormat.Html);
                richEdit.Document.AppendDocumentContent(documentServer.Document.Range);
            }
        }
        #endregion #LoadDocumentMethodCustomActionHandler


        static void IsSelectionInTable(RichEditControl richEditControl, BarButtonItem buttonCustomAction) {
            #region #IsSelectionInTable
            richEditControl.Tag = buttonCustomAction;
            richEditControl.SelectionChanged += richEditControl_SelectionChanged;
            richEditControl.LoadDocument("Documents\\TableSampleDocument.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml);
            #endregion #IsSelectionInTable
        }

        #region #IsSelectionInTableAdditionalModules
        static void richEditControl_SelectionChanged(object sender, EventArgs e) {
            RichEditControl richEdit = (sender as RichEditControl);
            BarButtonItem barButton = richEdit.Tag as BarButtonItem;
            barButton.Enabled = !richEdit.IsSelectionInTable();
        }
        #endregion #IsSelectionInTableAdditionalModules


        static void GetPositionFromPoint(RichEditControl richEditControl, BarButtonItem buttonCustomAction) {
            #region #GetPositionFromPoint
            richEditControl.MouseMove += richEditControl1_MouseMove;
            richEditControl.LoadDocument("Documents\\LoremDocumentTest.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml);
            richEditControl.Document.InsertText(richEditControl.Document.Range.Start, "Hovering a mouse in this document results in showing a tooltip with a current document position and a character in this position\r\n\r\n");
            #endregion #GetPositionFromPoint
        }


        #region #GetPositionFromPointAdditionalModules
        static DevExpress.Utils.ToolTipController testToolTipController = new DevExpress.Utils.ToolTipController();

        static void richEditControl1_MouseMove(object sender, System.Windows.Forms.MouseEventArgs e) {
            RichEditControl richEdit = sender as RichEditControl;
            System.Drawing.Point docPoint = 
                DevExpress.Office.Utils.Units.PixelsToDocuments(e.Location, richEdit.DpiX, richEdit.DpiY);

            DocumentPosition pos = richEdit.GetPositionFromPoint(docPoint);
            if(pos == null) return;
            string currentToolTipText = String.Format("Position: {0}, Character: {1}", pos.ToString(), 
                richEdit.Document.GetText(richEdit.Document.CreateRange(pos, 1)));

            DevExpress.Utils.ToolTipControlInfo info = 
                new DevExpress.Utils.ToolTipControlInfo(currentToolTipText, currentToolTipText);
            info.ToolTipPosition = System.Windows.Forms.Form.MousePosition;
            testToolTipController.ShowHint(info);
        }
        #endregion #GetPositionFromPointAdditionalModules

        static void GetBoundsFromDocumentPosition(RichEditControl richEditControl, BarButtonItem buttonCustomAction) {
            #region #GetBoundsFromDocumentPosition
            richEditControl.Paint += richEditControl_Paint;
            richEditControl.LoadDocument("Documents\\LoremDocumentTest.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml);
            richEditControl.Document.InsertText(richEditControl.Document.Range.Start, "A document row which contains a caret position is highlighted\r\n\r\n");
            #endregion #GetBoundsFromDocumentPosition
        }

        #region #GetBoundsFromDocumentPositionAdditionalModules
        static void richEditControl_Paint(object sender, System.Windows.Forms.PaintEventArgs e) {
            RichEditControl richEdit = sender as RichEditControl;
            DocumentPosition pos = richEdit.Document.CaretPosition;
            if(pos != null) {
                System.Drawing.Rectangle rect = DevExpress.Office.Utils.Units.DocumentsToPixels(
                    richEdit.GetBoundsFromPosition(pos), 
                    richEdit.DpiX, 
                    richEdit.DpiY);

                Section firstSection = richEdit.Document.Sections[0];
                int pageWidth = Convert.ToInt32(firstSection.Page.Width - firstSection.Margins.Left - firstSection.Margins.Right);
                e.Graphics.DrawLine(System.Drawing.Pens.Red, 
                    new System.Drawing.Point(0, rect.Bottom), 
                    new System.Drawing.Point(pageWidth, rect.Bottom));
            }
        }
        #endregion #GetBoundsFromDocumentPositionAdditionalModules


        static void ExportDocumentToPdfFormat(RichEditControl richEditControl, BarButtonItem buttonCustomAction) {
            #region #ExportDocumentToPdfFormat
            buttonCustomAction.ItemClick += buttonCustomAction_ItemClick_PDF;
            buttonCustomAction.Tag = richEditControl;
            richEditControl.Text = "Click the 'Custom Action' button located in the ribbon ";
            richEditControl.Text += "to load the 'Grimm.docx' file into RichEditControl's document and export the document into PDF format\r\n";
            #endregion #ExportDocumentToPdfFormat
        }

        #region #ExportDocumentToPdfFormatCustomActionHandler
        static void buttonCustomAction_ItemClick_PDF(object sender, ItemClickEventArgs e) {
            RichEditControl richEdit = e.Item.Tag as RichEditControl;
            richEdit.LoadDocument("Documents\\Grimm.docx");
            //Set the required export options:
            DevExpress.XtraPrinting.PdfExportOptions options = new DevExpress.XtraPrinting.PdfExportOptions();
            options.DocumentOptions.Author = "Mark Jones";
            options.Compressed = false;
            options.ImageQuality = DevExpress.XtraPrinting.PdfJpegImageQuality.High;
            //Export the document to the file:
            richEdit.ExportToPdf("resultingDocument.pdf", options);
            //Export the document to the file stream:
            using (FileStream pdfFileStream = new FileStream("resultingDocumentFromStream.pdf", FileMode.Create)) {
                richEdit.ExportToPdf(pdfFileStream, options);
            }

            System.Diagnostics.Process.Start("resultingDocument.pdf");
        }
        #endregion #ExportDocumentToPdfFormatCustomActionHandler

        static void ExportDocumentToHtmlFormat(RichEditControl richEditControl, BarButtonItem buttonCustomAction) {
            #region #ExportDocumentToHtmlFormat
            buttonCustomAction.ItemClick += buttonCustomAction_ItemClick_Html;
            buttonCustomAction.Tag = richEditControl;
            richEditControl.Text = "Click the 'Custom Action' button located in the ribbon ";
            richEditControl.Text += "to load the 'Grimm.docx' file into RichEditControl's document and export the document into HTML format\r\n";
            #endregion #ExportDocumentToHtmlFormat
        }

        #region #ExportDocumentToHtmlFormatCustomActionHandler
        static void buttonCustomAction_ItemClick_Html(object sender, ItemClickEventArgs e) {
            RichEditControl richEdit = e.Item.Tag as RichEditControl;
            richEdit.LoadDocument("Documents\\Grimm.docx");
            //Export document to the file:
            richEdit.SaveDocument("resultingDocument.html", DocumentFormat.Html);
            //Export document to the stream:
            using (FileStream htmlFileStream = new FileStream("Document_HTML.html", FileMode.Create)) {
                richEdit.SaveDocument(htmlFileStream, DocumentFormat.Html);
            }

            System.Diagnostics.Process.Start("resultingDocument.html");
        }
        #endregion #ExportDocumentToHtmlFormatCustomActionHandler

        static void DeselectAllTextInDocument(RichEditControl richEditControl, BarButtonItem buttonCustomAction) {
            #region #DeselectAllTextInDocument
            buttonCustomAction.ItemClick += buttonCustomAction_ItemClick_Deselect;
            buttonCustomAction.Tag = richEditControl;
            richEditControl.Text = "Select text in a document and сlick the 'Custom Action' button located in the ribbon.\r\n";
            richEditControl.Text += "As a result, the previously selected text is deselcted and the caret position is set to the end of the text\r\n";
            #endregion #DeselectAllTextInDocument
        }

        #region #DeselectAllTextInDocumentCustomActionHandler
        static void buttonCustomAction_ItemClick_Deselect(object sender, ItemClickEventArgs e) {
            RichEditControl richEdit = e.Item.Tag as RichEditControl;
            DocumentPosition endOfSelection = richEdit.Document.Selection.End;
            richEdit.DeselectAll();
            richEdit.Document.CaretPosition = endOfSelection;
        }
        #endregion #DeselectAllTextInDocumentCustomActionHandler

        static void CutSelectedText(RichEditControl richEditControl, BarButtonItem buttonCustomAction) {
            #region #CutSelectedText
            buttonCustomAction.ItemClick += buttonCustomAction_ItemClick_Cut;
            buttonCustomAction.Tag = richEditControl;
            richEditControl.Text = "Click the 'Custom Action' button located in the ribbon ";
            richEditControl.Text += "to cut the 'Custom Action' text block \r\n";
            #endregion #CutSelectedText
        }

        #region #CutSelectedTextCustomActionHandler
        static void buttonCustomAction_ItemClick_Cut(object sender, ItemClickEventArgs e) {
            RichEditControl richEdit = e.Item.Tag as RichEditControl;
            DocumentRange[] foundRanges = richEdit.Document.FindAll("Custom Action", SearchOptions.None);
            for(int i = 0; i < foundRanges.Length; i++) {
                richEdit.Document.Selection = foundRanges[i];
                richEdit.Cut();
            }
        }
        #endregion #CutSelectedTextCustomActionHandler

        static void ReplaceSelectedText(RichEditControl richEditControl, BarButtonItem buttonCustomAction) {
            #region #ReplaceSelectedText
            buttonCustomAction.ItemClick += buttonCustomAction_ItemClick_Replace;
            buttonCustomAction.Tag = richEditControl;
            richEditControl.Text = "Click the 'Custom Action' button located in the ribbon ";
            richEditControl.Text += "to replace selected text with asterisks.\r\n";
            #endregion #ReplaceSelectedText
        }

        #region #ReplaceSelectedTextCustomActionHandler
        static void buttonCustomAction_ItemClick_Replace(object sender, ItemClickEventArgs e) {
            RichEditControl richEdit = e.Item.Tag as RichEditControl;
            DocumentRange range = richEdit.Document.Selection;
            int selLength = range.Length;
            string s = new String('*', selLength);
            SubDocument doc = range.BeginUpdateDocument();
            doc.InsertSingleLineText(range.Start, s);
            DocumentRange rangeToRemove = doc.CreateRange(range.Start, selLength);
            doc.Delete(rangeToRemove);
            range.EndUpdateDocument(doc);
        }
        #endregion #ReplaceSelectedTextCustomActionHandler


        static void CreateNewDocument(RichEditControl richEditControl, BarButtonItem buttonCustomAction) {
            #region #CreateNewDocument
            buttonCustomAction.ItemClick += buttonCustomAction_ItemClick_NewDocument;
            buttonCustomAction.Tag = richEditControl;
            richEditControl.Text = "Click the 'Custom Action' button located in the ribbon ";
            richEditControl.Text += "to create a new document \r\n";
            #endregion #CreateNewDocument
        }

        #region #CreateNewDocumentCustomActionHandler
        static void buttonCustomAction_ItemClick_NewDocument(object sender, ItemClickEventArgs e) {
            RichEditControl richEdit = e.Item.Tag as RichEditControl;
            richEdit.CreateNewDocument();
        }
        #endregion #CreateNewDocumentCustomActionHandler

        static void CreateAndExecuteCommands(RichEditControl richEditControl, BarButtonItem buttonCustomAction) {
            #region #CreateAndExecuteCommands
            buttonCustomAction.ItemClick += buttonCustomAction_ItemClick_Commands;
            buttonCustomAction.Tag = richEditControl;
            richEditControl.Text = "Click the 'Custom Action' button located in the ribbon to capitalize each word in this text,\r\n";
            richEditControl.Text += "to change font of this text to Bold and show the Print Preview window with a result.\r\n";
            richEditControl.Text += "All these actions are performed using corresponding RichEditControl commands.\r\n";
            #endregion #CreateAndExecuteCommands
        }

        #region #CreateAndExecuteCommandsCustomActionHandler
        static void buttonCustomAction_ItemClick_Commands(object sender, ItemClickEventArgs e) {
            RichEditControl richEdit = e.Item.Tag as RichEditControl;
            richEdit.SelectAll();

            RichEditCommand capCommand = richEdit.CreateCommand(RichEditCommandId.CapitalizeEachWordTextCase);
            capCommand.ForceExecute(capCommand.CreateDefaultCommandUIState());

            RichEditCommand boldCommand = richEdit.CreateCommand(RichEditCommandId.ToggleFontBold);
            boldCommand.ForceExecute(boldCommand.CreateDefaultCommandUIState());

            RichEditCommand changeFontColorCommand = richEdit.CreateCommand(RichEditCommandId.ChangeFontBackColor);
            DevExpress.Utils.Commands.ICommandUIState state = changeFontColorCommand.CreateDefaultCommandUIState();
            state.EditValue = Color.Yellow;
            changeFontColorCommand.ForceExecute(state);

            RichEditCommand previewCommand = richEdit.CreateCommand(RichEditCommandId.PrintPreview);
            previewCommand.ForceExecute(previewCommand.CreateDefaultCommandUIState());

            richEdit.DeselectAll();
        }
        #endregion #CreateAndExecuteCommandsCustomActionHandler

        static void CopySelectedText(RichEditControl richEditControl, BarButtonItem buttonCustomAction) {
            #region #CopySelectedText
            richEditControl.Text = "Open any text editor (Notepad, Word) and press the 'Ctrl+V' shortcut, the 'text from RichEditControl' string will be pasted into the editor's document";
            DocumentRange[] foundRanges = richEditControl.Document.FindAll("text from RichEditControl", SearchOptions.None);
            richEditControl.Document.Selection = foundRanges[0];
            richEditControl.Copy();
            #endregion #CopySelectedText
        }

        static void CustomizePredefinedShortcutKeys(RichEditControl richEditControl, BarButtonItem buttonCustomAction) {
            #region #CustomizePredefinedShortcutKeys
            richEditControl.Text = "Use 'Ctrl+G' shortcut to show/hide whitespace characters";
            richEditControl.Text += "\r\nA new document can not be created by pressing 'Ctrl+N' shortcut since this shortcut is disabled";
            richEditControl.AssignShortcutKeyToCommand(System.Windows.Forms.Keys.Control, System.Windows.Forms.Keys.G, RichEditCommandId.ToggleShowWhitespace);
            richEditControl.RemoveShortcutKey(System.Windows.Forms.Keys.Control, System.Windows.Forms.Keys.N);
            #endregion #CustomizePredefinedShortcutKeys
        }

        static void ReplacePredefinedService(RichEditControl richEditControl, BarButtonItem buttonCustomAction) {
            richEditControl.ActiveViewType = RichEditViewType.PrintLayout;
            #region #ReplacePredefinedService
            richEditControl.Text = "A message box is displayed after saving a document using the 'Save' or 'Save As' button click since custom commands.";
            var myCommandFactory = new CustomRichEditCommandFactoryService(richEditControl, richEditControl.GetService<IRichEditCommandFactoryService>());
            richEditControl.ReplaceService<IRichEditCommandFactoryService>(myCommandFactory);
            #endregion #ReplacePredefinedService
        }

        #region #ReplacePredefinedServiceAdditionalModules
        public class CustomRichEditCommandFactoryService : IRichEditCommandFactoryService {
            readonly IRichEditCommandFactoryService service;
            readonly RichEditControl control;

            public CustomRichEditCommandFactoryService(RichEditControl control, IRichEditCommandFactoryService service) {
                DevExpress.Utils.Guard.ArgumentNotNull(control, "control");
                DevExpress.Utils.Guard.ArgumentNotNull(service, "service");
                this.control = control;
                this.service = service;
            }

            public RichEditCommand CreateCommand(RichEditCommandId id) {
                if(id == RichEditCommandId.FileSaveAs) {
                    return new CustomSaveDocumentAsCommand(control);
                }
                if(id == RichEditCommandId.FileSave) {
                    return new CustomSaveDocumentCommand(control);
                }
                return service.CreateCommand(id);
            }
        }

        public class CustomSaveDocumentCommand : SaveDocumentCommand {
            public CustomSaveDocumentCommand(IRichEditControl richEdit) : base(richEdit) { }

            protected override void ExecuteCore() {
                base.ExecuteCore();
                if(!(this.Control as RichEditControl).Modified) {
                    MessageBox.Show("Document is saved successfully");
                }
            }
        }


        public class CustomSaveDocumentAsCommand : SaveDocumentAsCommand {
            public CustomSaveDocumentAsCommand(IRichEditControl richEdit) : base(richEdit) { }

            protected override void ExecuteCore() {
                (this.Control as RichEditControl).Modified = true;
                base.ExecuteCore();
                if(!(this.Control as RichEditControl).Modified) {
                    MessageBox.Show("Document is saved successfully");    
                }                
            }
        }
        #endregion #ReplacePredefinedServiceAdditionalModules

        static void SimpleViewLineNumbering(RichEditControl richEditControl, BarButtonItem buttonCustomAction) {
            #region #SimpleViewLineNumbering
            richEditControl.LoadDocument("Documents\\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml);
            richEditControl.Views.SimpleView.Padding = new DevExpress.Portable.PortablePadding(60, 4, 4, 0);
            richEditControl.Views.SimpleView.AllowDisplayLineNumbers = true;
            richEditControl.ActiveViewType = RichEditViewType.Simple;
            richEditControl.Document.Sections[0].LineNumbering.CountBy = 1;
            richEditControl.Document.CharacterStyles["Line Number"].ForeColor = Color.LightGray;
            richEditControl.Document.CharacterStyles["Line Number"].Bold = true;
            #endregion #SimpleViewLineNumbering
        }

    }
}
