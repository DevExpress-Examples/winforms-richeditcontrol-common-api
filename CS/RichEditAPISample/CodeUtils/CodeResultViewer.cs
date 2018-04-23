using System;
using System.CodeDom.Compiler;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.CSharp;
using System.Globalization;
using DevExpress.Spreadsheet;

namespace RichEditAPISample
{
    public abstract class RichEditExampleCodeEvaluator : ExampleCodeEvaluator
    {
        protected override string[] GetModuleAssemblies()
        {
            return new string[] { 
                AssemblyInfo.SRAssemblyRichEditCore,
                AssemblyInfo.SRAssemblyRichEdit,
                AssemblyInfo.SRAssemblyBars,
                AssemblyInfo.SRAssemblyEditors
            };
        }
        protected override string GetExampleClassName()
        {
            return "RichEditCodeResultViewer.ExampleItem";
        }
    }
    #region RichEditCSExampleCodeEvaluator
    public class RichEditCSExampleCodeEvaluator : RichEditExampleCodeEvaluator
    {

        protected override CodeDomProvider GetCodeDomProvider()
        {
            return new CSharpCodeProvider();
        }
        const string codeStart =
      "using System;\r\n" +
      "using DevExpress.XtraRichEdit;\r\n" +
      "using DevExpress.XtraRichEdit.API.Native;\r\n" +
      "using DevExpress.XtraRichEdit.Services;\r\n" +
      "using DevExpress.XtraRichEdit.Commands;\r\n" +
      "using DevExpress.XtraBars;\r\n" +
      "using System.Drawing;\r\n" +
      "using System.Windows.Forms;\r\n" +
      "using DevExpress.Utils;\r\n" +
      "using System.IO;\r\n" +
      "using System.Diagnostics;\r\n" +
      "using System.Xml;\r\n" +
      "using System.Data;\r\n" +
      "using System.Collections.Generic;\r\n" +
      "using System.Linq;\r\n" +
      "using System.Globalization;\r\n" +
      "using System.Reflection;\r\n" +
      "using System.ComponentModel;\r\n" +
      "namespace RichEditCodeResultViewer { \r\n" +
      "public class ExampleItem { \r\n" +
      "        public static void Process(RichEditControl richEditControl, BarButtonItem buttonCustomAction) { \r\n" +
      "                FieldInfo fItemClick = typeof(BarItem).GetField(\"itemClick\", BindingFlags.Static | BindingFlags.NonPublic);\r\n" +
      "                object objItemClick = fItemClick.GetValue(buttonCustomAction);\r\n" +
      "                PropertyInfo piItemClick = buttonCustomAction.GetType().GetProperty(\"Events\", BindingFlags.NonPublic | BindingFlags.Instance);\r\n" +
      "                EventHandlerList listItemClick = (EventHandlerList)piItemClick.GetValue(buttonCustomAction, null);\r\n" +
      "                listItemClick.RemoveHandler(objItemClick, listItemClick[objItemClick]);\r\n" +
      "\r\n" +
      "                FieldInfo fEventPaint = typeof(Control).GetField(\"EventPaint\", BindingFlags.Static | BindingFlags.NonPublic);\r\n" +
      "                object objEventPaint = fEventPaint.GetValue(richEditControl);\r\n" +
      "                PropertyInfo piEventPaint = richEditControl.GetType().GetProperty(\"Events\", BindingFlags.NonPublic | BindingFlags.Instance);\r\n" +
      "                EventHandlerList listEventPaint = (EventHandlerList)piEventPaint.GetValue(richEditControl, null);\r\n" +
      "                listEventPaint.RemoveHandler(objEventPaint, listEventPaint[objEventPaint]);\r\n" +
      "\r\n" +
      "                FieldInfo fEventMouseHover = typeof(Control).GetField(\"EventMouseMove\", BindingFlags.Static | BindingFlags.NonPublic);\r\n" +
      "                object objEventMouseHover = fEventMouseHover.GetValue(richEditControl);\r\n" +
      "                PropertyInfo piMouseHover = typeof(Control).GetProperty(\"Events\", BindingFlags.NonPublic | BindingFlags.Instance);\r\n" +
      "                EventHandlerList listEventMouseHover = (EventHandlerList)piMouseHover.GetValue(richEditControl, null);\r\n" +
      "                listEventMouseHover.RemoveHandler(objEventMouseHover, listEventMouseHover[objEventMouseHover]);\r\n" +
      "\r\n";

        const string mainMethodEnd =
        "       \r\n }\r\n";
        
        const string codeEnd =
            "    }\r\n" +
            "}\r\n";
        protected override string CodeStart { get { return codeStart; } }
        protected override string MainMethodEnd { get { return mainMethodEnd; } }
        protected override string CodeEnd { get { return codeEnd; } }
    }
    #endregion
    #region RichEditVbExampleCodeEvaluator
    public class RichEditVbExampleCodeEvaluator : RichEditExampleCodeEvaluator
    {

        protected override CodeDomProvider GetCodeDomProvider()
        {
            return new Microsoft.VisualBasic.VBCodeProvider();
        }
        const string codeStart =
      "Imports Microsoft.VisualBasic\r\n" +
      "Imports System\r\n" +
      "Imports DevExpress.XtraRichEdit\r\n" +
      "Imports DevExpress.XtraRichEdit.API.Native\r\n" +
      "Imports DevExpress.XtraRichEdit.Services\r\n" +
      "Imports DevExpress.XtraRichEdit.Commands\r\n" +
      "Imports DevExpress.XtraBars\r\n" +
      "Imports System.Drawing\r\n" +
      "Imports System.Windows.Forms\r\n" +
      "Imports DevExpress.Utils\r\n" +
      "Imports System.IO\r\n" +
      "Imports System.Diagnostics\r\n" +
      "Imports System.Xml\r\n" +
      "Imports System.Data\r\n" +
      "Imports System.Collections.Generic\r\n" +
      "Imports System.Linq;\r\n" +
      "Imports System.Globalization\r\n" +
      "Imports System.Reflection;\r\n" +
      "Imports System.ComponentModel;\r\n" +
      "Namespace RichEditCodeResultViewer\r\n" +
      "	Public Class ExampleItem\r\n" +
      "		Public Shared Sub Process(ByVal richEditControl As RichEditControl, ByVal buttonCustomAction As BarButtonItem)\r\n" +
      "                Dim fItemClick As FieldInfo = GetType(BarItem).GetField(\"itemClick\", BindingFlags.Static Or BindingFlags.NonPublic)\r\n" +
      "                Dim objItemClick As Object = fItemClick.GetValue(buttonCustomAction)\r\n" +
      "                Dim piItemClick As PropertyInfo = buttonCustomAction.GetType().GetProperty(\"Events\", BindingFlags.NonPublic Or BindingFlags.Instance)\r\n" +
      "                Dim listItemClick As EventHandlerList = CType(piItemClick.GetValue(buttonCustomAction, Nothing), EventHandlerList)\r\n" +
      "                listItemClick.RemoveHandler(objItemClick, listItemClick(objItemClick))\r\n" +
      "\r\n" +
      "                Dim fEventPaint As FieldInfo = GetType(Control).GetField(\"EventPaint\", BindingFlags.Static Or BindingFlags.NonPublic)\r\n" +
      "                Dim objEventPaint As Object = fEventPaint.GetValue(richEditControl)\r\n" +
      "                Dim piEventPaint As PropertyInfo = richEditControl.GetType().GetProperty(\"Events\", BindingFlags.NonPublic Or BindingFlags.Instance)\r\n" +
      "                Dim listEventPaint As EventHandlerList = CType(piEventPaint.GetValue(richEditControl, Nothing), EventHandlerList)\r\n" +
      "                listEventPaint.RemoveHandler(objEventPaint, listEventPaint(objEventPaint))\r\n" +
      "\r\n" +
      "                Dim fEventMouseHover As FieldInfo = GetType(Control).GetField(\"EventMouseMove\", BindingFlags.Static Or BindingFlags.NonPublic)\r\n" +
      "                Dim objEventMouseHover As Object = fEventMouseHover.GetValue(richEditControl)\r\n" +
      "                Dim piMouseHover As PropertyInfo = GetType(Control).GetProperty(\"Events\", BindingFlags.NonPublic Or BindingFlags.Instance)\r\n" +
      "                Dim listEventMouseHover As EventHandlerList = CType(piMouseHover.GetValue(richEditControl, Nothing), EventHandlerList)\r\n" +
      "                listEventMouseHover.RemoveHandler(objEventMouseHover, listEventMouseHover(objEventMouseHover))\r\n" +
      "\r\n";

        const string mainMethodEnd =
        "\r\n		End Sub\r\n\r\n";

        const string codeEnd = 
        "	End Class\r\n" +
        "End Namespace\r\n";

        protected override string CodeStart { get { return codeStart; } }
        protected override string MainMethodEnd { get { return mainMethodEnd; } }
        protected override string CodeEnd { get { return codeEnd; } }
    }
    #endregion
}
