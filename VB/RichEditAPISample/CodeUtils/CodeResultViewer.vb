Imports Microsoft.VisualBasic
Imports System
Imports System.CodeDom.Compiler
Imports System.IO
Imports System.Reflection
Imports System.Windows.Forms
Imports Microsoft.CSharp
Imports System.Globalization
Imports DevExpress.Spreadsheet

Namespace RichEditAPISample
	Public MustInherit Class RichEditExampleCodeEvaluator
		Inherits ExampleCodeEvaluator
		Protected Overrides Function GetModuleAssemblies() As String()
			Return New String() { AssemblyInfo.SRAssemblyRichEditCore, AssemblyInfo.SRAssemblyRichEdit, AssemblyInfo.SRAssemblyBars, AssemblyInfo.SRAssemblyEditors }
		End Function
		Protected Overrides Function GetExampleClassName() As String
			Return "RichEditCodeResultViewer.ExampleItem"
		End Function
	End Class
	#Region "RichEditCSExampleCodeEvaluator"
	Public Class RichEditCSExampleCodeEvaluator
		Inherits RichEditExampleCodeEvaluator

		Protected Overrides Function GetCodeDomProvider() As CodeDomProvider
			Return New CSharpCodeProvider()
		End Function
        Private Const codeStart_Renamed As String = "using System;" & Constants.vbCrLf & "using DevExpress.XtraRichEdit;" & Constants.vbCrLf & "using DevExpress.XtraRichEdit.API.Native;" & Constants.vbCrLf & "using DevExpress.XtraRichEdit.Services;" & Constants.vbCrLf & "using DevExpress.XtraRichEdit.Commands;" & Constants.vbCrLf & "using DevExpress.XtraBars;" & Constants.vbCrLf & "using System.Drawing;" & Constants.vbCrLf & "using System.Windows.Forms;" & Constants.vbCrLf & "using DevExpress.Utils;" & Constants.vbCrLf & "using System.IO;" & Constants.vbCrLf & "using System.Diagnostics;" & Constants.vbCrLf & "using System.Xml;" & Constants.vbCrLf & "using System.Data;" & Constants.vbCrLf & "using System.Collections.Generic;" & Constants.vbCrLf & "using System.Linq;" & Constants.vbCrLf & "using System.Globalization;" & Constants.vbCrLf & "using System.Reflection;" & Constants.vbCrLf & "using System.ComponentModel;" & Constants.vbCrLf & "namespace RichEditCodeResultViewer { " & Constants.vbCrLf & "public class ExampleItem { " & Constants.vbCrLf & "        public static void Process(RichEditControl richEditControl, BarButtonItem buttonCustomAction) { " & Constants.vbCrLf & "                FieldInfo fItemClick = typeof(BarItem).GetField(""itemClick"", BindingFlags.Static | BindingFlags.NonPublic);" & Constants.vbCrLf & "                object objItemClick = fItemClick.GetValue(buttonCustomAction);" & Constants.vbCrLf & "                PropertyInfo piItemClick = buttonCustomAction.GetType().GetProperty(""Events"", BindingFlags.NonPublic | BindingFlags.Instance);" & Constants.vbCrLf & "                EventHandlerList listItemClick = (EventHandlerList)piItemClick.GetValue(buttonCustomAction, null);" & Constants.vbCrLf & "                listItemClick.RemoveHandler(objItemClick, listItemClick[objItemClick]);" & Constants.vbCrLf & Constants.vbCrLf & "                FieldInfo fEventPaint = typeof(Control).GetField(""EventPaint"", BindingFlags.Static | BindingFlags.NonPublic);" & Constants.vbCrLf & "                object objEventPaint = fEventPaint.GetValue(richEditControl);" & Constants.vbCrLf & "                PropertyInfo piEventPaint = richEditControl.GetType().GetProperty(""Events"", BindingFlags.NonPublic | BindingFlags.Instance);" & Constants.vbCrLf & "                EventHandlerList listEventPaint = (EventHandlerList)piEventPaint.GetValue(richEditControl, null);" & Constants.vbCrLf & "                listEventPaint.RemoveHandler(objEventPaint, listEventPaint[objEventPaint]);" & Constants.vbCrLf & Constants.vbCrLf & "                FieldInfo fEventMouseHover = typeof(Control).GetField(""EventMouseMove"", BindingFlags.Static | BindingFlags.NonPublic);" & Constants.vbCrLf & "                object objEventMouseHover = fEventMouseHover.GetValue(richEditControl);" & Constants.vbCrLf & "                PropertyInfo piMouseHover = typeof(Control).GetProperty(""Events"", BindingFlags.NonPublic | BindingFlags.Instance);" & Constants.vbCrLf & "                EventHandlerList listEventMouseHover = (EventHandlerList)piMouseHover.GetValue(richEditControl, null);" & Constants.vbCrLf & "                listEventMouseHover.RemoveHandler(objEventMouseHover, listEventMouseHover[objEventMouseHover]);" & Constants.vbCrLf & Constants.vbCrLf

		Private Const mainMethodEnd_Renamed As String = "       " & Constants.vbCrLf & " }" & Constants.vbCrLf

		Private Const codeEnd_Renamed As String = "    }" & Constants.vbCrLf & "}" & Constants.vbCrLf
		Protected Overrides ReadOnly Property CodeStart() As String
			Get
				Return codeStart_Renamed
			End Get
		End Property
		Protected Overrides ReadOnly Property MainMethodEnd() As String
			Get
				Return mainMethodEnd_Renamed
			End Get
		End Property
		Protected Overrides ReadOnly Property CodeEnd() As String
			Get
				Return codeEnd_Renamed
			End Get
		End Property
	End Class
	#End Region
	#Region "RichEditVbExampleCodeEvaluator"
	Public Class RichEditVbExampleCodeEvaluator
		Inherits RichEditExampleCodeEvaluator

		Protected Overrides Function GetCodeDomProvider() As CodeDomProvider
			Return New Microsoft.VisualBasic.VBCodeProvider()
		End Function
        Private Const codeStart_Renamed As String = "Imports Microsoft.VisualBasic" & Constants.vbCrLf & "Imports System" & Constants.vbCrLf & "Imports DevExpress.XtraRichEdit" & Constants.vbCrLf & "Imports DevExpress.XtraRichEdit.API.Native" & Constants.vbCrLf & "Imports DevExpress.XtraRichEdit.Services" & Constants.vbCrLf & "Imports DevExpress.XtraRichEdit.Commands" & Constants.vbCrLf & "Imports DevExpress.XtraBars" & Constants.vbCrLf & "Imports System.Drawing" & Constants.vbCrLf & "Imports System.Windows.Forms" & Constants.vbCrLf & "Imports DevExpress.Utils" & Constants.vbCrLf & "Imports System.IO" & Constants.vbCrLf & "Imports System.Diagnostics" & Constants.vbCrLf & "Imports System.Xml" & Constants.vbCrLf & "Imports System.Data" & Constants.vbCrLf & "Imports System.Collections.Generic" & Constants.vbCrLf & "Imports System.Linq" & Constants.vbCrLf & "Imports System.Globalization" & Constants.vbCrLf & "Imports System.Reflection" & Constants.vbCrLf & "Imports System.ComponentModel" & Constants.vbCrLf & "Namespace RichEditCodeResultViewer" & Constants.vbCrLf & "	Public Class ExampleItem" & Constants.vbCrLf & "		Public Shared Sub Process(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)" & Constants.vbCrLf & "                Dim fItemClick As FieldInfo = GetType(BarItem).GetField(""itemClick"", BindingFlags.Static Or BindingFlags.NonPublic)" & Constants.vbCrLf & "                Dim objItemClick As Object = fItemClick.GetValue(buttonCustomAction)" & Constants.vbCrLf & "                Dim piItemClick As PropertyInfo = buttonCustomAction.GetType().GetProperty(""Events"", BindingFlags.NonPublic Or BindingFlags.Instance)" & Constants.vbCrLf & "                Dim listItemClick As EventHandlerList = CType(piItemClick.GetValue(buttonCustomAction, Nothing), EventHandlerList)" & Constants.vbCrLf & "                listItemClick.RemoveHandler(objItemClick, listItemClick(objItemClick))" & Constants.vbCrLf & Constants.vbCrLf & "                Dim fEventPaint As FieldInfo = GetType(Control).GetField(""EventPaint"", BindingFlags.Static Or BindingFlags.NonPublic)" & Constants.vbCrLf & "                Dim objEventPaint As Object = fEventPaint.GetValue(richEditControl)" & Constants.vbCrLf & "                Dim piEventPaint As PropertyInfo = richEditControl.GetType().GetProperty(""Events"", BindingFlags.NonPublic Or BindingFlags.Instance)" & Constants.vbCrLf & "                Dim listEventPaint As EventHandlerList = CType(piEventPaint.GetValue(richEditControl, Nothing), EventHandlerList)" & Constants.vbCrLf & "                listEventPaint.RemoveHandler(objEventPaint, listEventPaint(objEventPaint))" & Constants.vbCrLf & Constants.vbCrLf & "                Dim fEventMouseHover As FieldInfo = GetType(Control).GetField(""EventMouseMove"", BindingFlags.Static Or BindingFlags.NonPublic)" & Constants.vbCrLf & "                Dim objEventMouseHover As Object = fEventMouseHover.GetValue(richEditControl)" & Constants.vbCrLf & "                Dim piMouseHover As PropertyInfo = GetType(Control).GetProperty(""Events"", BindingFlags.NonPublic Or BindingFlags.Instance)" & Constants.vbCrLf & "                Dim listEventMouseHover As EventHandlerList = CType(piMouseHover.GetValue(richEditControl, Nothing), EventHandlerList)" & Constants.vbCrLf & "                listEventMouseHover.RemoveHandler(objEventMouseHover, listEventMouseHover(objEventMouseHover))" & Constants.vbCrLf & Constants.vbCrLf

        Private Const mainMethodEnd_Renamed As String = Constants.vbCrLf & "		End Sub" & Constants.vbCrLf

        Private Const codeEnd_Renamed As String = Constants.vbCrLf & "	End Class" & Constants.vbCrLf & "End Namespace" & Constants.vbCrLf

		Protected Overrides ReadOnly Property CodeStart() As String
			Get
				Return codeStart_Renamed
			End Get
		End Property
		Protected Overrides ReadOnly Property MainMethodEnd() As String
			Get
				Return mainMethodEnd_Renamed
			End Get
		End Property
		Protected Overrides ReadOnly Property CodeEnd() As String
			Get
				Return codeEnd_Renamed
			End Get
		End Property
	End Class
	#End Region
End Namespace
