Imports Microsoft.VisualBasic
Imports System.CodeDom.Compiler
Imports DevExpress.Internal
Imports System.Reflection
Imports System.IO
Imports System.Text.RegularExpressions
Imports System
Imports System.Collections.Generic
Imports System.Text

Namespace RichEditAPISample
	Public MustInherit Class ExampleCodeEvaluator
		Protected MustOverride ReadOnly Property CodeStart() As String
		Protected MustOverride ReadOnly Property MainMethodEnd() As String
		Protected MustOverride ReadOnly Property CodeEnd() As String
		Protected MustOverride Function GetCodeDomProvider() As CodeDomProvider
		Protected MustOverride Function GetModuleAssemblies() As String()
		Protected MustOverride Function GetExampleClassName() As String

		Public Function ExcecuteCodeAndGenerateDocument(ByVal args As CodeEvaluationEventArgs) As Boolean
			Dim theCode As String = String.Concat(CodeStart, args.Code, MainMethodEnd, args.CustomActionHandler, args.AdditionalModules, CodeEnd)
			Dim linesOfCode() As String = { theCode }
			Return CompileAndRun(linesOfCode, args.EvaluationParameter, args.AdditionalParameter)
		End Function

		Protected Friend Function CompileAndRun(ByVal linesOfCode() As String, ByVal evaluationParameter As Object, ByVal additionalParameter As Object) As Boolean
			Dim CompilerParams As New CompilerParameters()

			CompilerParams.GenerateInMemory = True
			CompilerParams.TreatWarningsAsErrors = False
			CompilerParams.GenerateExecutable = False

			Dim referencesSystem() As String = { "System.dll", "System.Windows.Forms.dll", "System.Data.dll", "System.Xml.dll", "System.Core.dll", "System.Drawing.dll" }

			Dim referencesDX() As String = { AssemblyInfo.SRAssemblyData, AssemblyInfo.SRAssemblyOfficeCore, AssemblyInfo.SRAssemblyPrintingCore, AssemblyInfo.SRAssemblyPrinting, AssemblyInfo.SRAssemblyDocs, AssemblyInfo.SRAssemblyUtils }

			Dim referencesDXModule() As String = GetModuleAssemblies()

			Dim references(referencesSystem.Length + referencesDX.Length + referencesDXModule.Length - 1) As String

			For referenceIndex As Integer = 0 To referencesSystem.Length - 1
				references(referenceIndex) = referencesSystem(referenceIndex)
			Next referenceIndex

			Dim i As Integer = 0
			Dim initial As Integer = referencesSystem.Length
			Do While i < referencesDX.Length
				Dim [assembly] As System.Reflection.Assembly = System.Reflection.Assembly.Load(referencesDX(i) + AssemblyInfo.FullAssemblyVersionExtension)
				If [assembly] IsNot Nothing Then
					references(i + initial) = [assembly].Location
				End If
				i += 1
			Loop

			i = 0
			initial = referencesSystem.Length + referencesDX.Length
			Do While i < referencesDXModule.Length
				Dim [assembly] As System.Reflection.Assembly = System.Reflection.Assembly.Load(referencesDXModule(i) + AssemblyInfo.FullAssemblyVersionExtension)
				If [assembly] IsNot Nothing Then
					references(i + initial) = [assembly].Location
				End If
				i += 1
			Loop

			CompilerParams.ReferencedAssemblies.AddRange(references)


			Dim provider As CodeDomProvider = GetCodeDomProvider()
			Dim compile As CompilerResults = provider.CompileAssemblyFromSource(CompilerParams, linesOfCode)

			If compile.Errors.HasErrors Then
				'string text = "Compile error: ";
				'foreach(CompilerError ce in compile.Errors) {
				'    text += "rn" + ce.ToString();
				'}
				'MessageBox.Show(text);
				Return False
			End If

			Dim [module] As System.Reflection.Module = Nothing
			Try
				[module] = compile.CompiledAssembly.GetModules()(0)
			Catch
			End Try
			Dim moduleType As Type = Nothing
			If [module] Is Nothing Then
				Return False
			End If
			moduleType = [module].GetType(GetExampleClassName())

			Dim methInfo As MethodInfo = Nothing
			If moduleType Is Nothing Then
				Return False
			End If
			methInfo = moduleType.GetMethod("Process")

			If methInfo IsNot Nothing Then
				Try
					methInfo.Invoke(Nothing, New Object() { evaluationParameter, additionalParameter })
				Catch e1 As Exception
					Return False ' an error
				End Try
				Return True
			End If
			Return False
		End Function
	End Class

	Public Class CodeExampleGroup
		Public Sub New()
		End Sub
		Private privateName As String
		Public Property Name() As String
			Get
				Return privateName
			End Get
			Set(ByVal value As String)
				privateName = value
			End Set
		End Property
		Private privateExamples As List(Of CodeExample)
		Public Property Examples() As List(Of CodeExample)
			Get
				Return privateExamples
			End Get
			Set(ByVal value As List(Of CodeExample))
				privateExamples = value
			End Set
		End Property
		Private privateId As Integer
		Public Property Id() As Integer
			Get
				Return privateId
			End Get
			Set(ByVal value As Integer)
				privateId = value
			End Set
		End Property
	End Class

	Public Class CodeExample
		Private privateCodeCS As String
		Public Property CodeCS() As String
			Get
				Return privateCodeCS
			End Get
			Set(ByVal value As String)
				privateCodeCS = value
			End Set
		End Property
		Private privateCodeVB As String
		Public Property CodeVB() As String
			Get
				Return privateCodeVB
			End Get
			Set(ByVal value As String)
				privateCodeVB = value
			End Set
		End Property
		Private privateAdditionalModulesCS As String
		Public Property AdditionalModulesCS() As String
			Get
				Return privateAdditionalModulesCS
			End Get
			Set(ByVal value As String)
				privateAdditionalModulesCS = value
			End Set
		End Property
		Private privateAdditionalModulesVB As String
		Public Property AdditionalModulesVB() As String
			Get
				Return privateAdditionalModulesVB
			End Get
			Set(ByVal value As String)
				privateAdditionalModulesVB = value
			End Set
		End Property
		Private privateRegionName As String
		Public Property RegionName() As String
			Get
				Return privateRegionName
			End Get
			Set(ByVal value As String)
				privateRegionName = value
			End Set
		End Property
		Private privateHumanReadableGroupName As String
		Public Property HumanReadableGroupName() As String
			Get
				Return privateHumanReadableGroupName
			End Get
			Set(ByVal value As String)
				privateHumanReadableGroupName = value
			End Set
		End Property
		Private privateExampleGroup As String
		Public Property ExampleGroup() As String
			Get
				Return privateExampleGroup
			End Get
			Set(ByVal value As String)
				privateExampleGroup = value
			End Set
		End Property
		Private privateHasCustomAction As Boolean
		Public Property HasCustomAction() As Boolean
			Get
				Return privateHasCustomAction
			End Get
			Set(ByVal value As Boolean)
				privateHasCustomAction = value
			End Set
		End Property
		Private privateCustomActionHandlerCS As String
		Public Property CustomActionHandlerCS() As String
			Get
				Return privateCustomActionHandlerCS
			End Get
			Set(ByVal value As String)
				privateCustomActionHandlerCS = value
			End Set
		End Property
		Private privateCustomActionHandlerVB As String
		Public Property CustomActionHandlerVB() As String
			Get
				Return privateCustomActionHandlerVB
			End Get
			Set(ByVal value As String)
				privateCustomActionHandlerVB = value
			End Set
		End Property
		Private privateId As Integer
		Public Property Id() As Integer
			Get
				Return privateId
			End Get
			Set(ByVal value As Integer)
				privateId = value
			End Set
		End Property
	End Class

	Public Enum ExampleLanguage
		Csharp = 0
		VB = 1
	End Enum

	#Region "CodeExampleDemoUtils"
	Public NotInheritable Class CodeExampleDemoUtils
		Private Sub New()
		End Sub
		Public Shared Function GatherExamplesFromProject(ByVal examplesPath As String, ByVal language As ExampleLanguage) As Dictionary(Of String, FileInfo)
			Dim result As New Dictionary(Of String, FileInfo)()
			For Each fileName As String In Directory.GetFiles(examplesPath, "*" & GetCodeExampleFileExtension(language))
				result.Add(Path.GetFileNameWithoutExtension(fileName), New FileInfo(fileName))
			Next fileName
			Return result
		End Function
		Public Shared Function GetCodeExampleFileExtension(ByVal language As ExampleLanguage) As String
			If language = ExampleLanguage.VB Then
				Return ".vb"
			End If
			Return ".cs"
		End Function
		Public Shared Function DeleteLeadingWhiteSpaces(ByVal lines() As String, ByVal stringToDelete As String) As String()
			Dim result(lines.Length - 1) As String
			Dim stringToDeleteLength As Integer = stringToDelete.Length

			For i As Integer = 0 To lines.Length - 1
				Dim index As Integer = lines(i).IndexOf(stringToDelete)
				result(i) = If((index >= 0), lines(i).Substring(index + stringToDeleteLength), lines(i))
			Next i
			Return result
		End Function
		Public Shared Function ConvertStringToMoreHumanReadableForm(ByVal exampleName As String) As String
			Dim result As String = SplitCamelCase(exampleName)
			result = result.Replace(" In ", " in ")
			result = result.Replace(" And ", " and ")
			result = result.Replace(" To ", " to ")
			result = result.Replace(" From ", " from ")
			result = result.Replace(" With ", " with ")
			result = result.Replace(" By ", " by ")
			Return result
		End Function
		Private Shared Function SplitCamelCase(ByVal exampleName As String) As String
			Dim length As Integer = exampleName.Length
			If length = 1 Then
				Return exampleName
			End If

			Dim result As New StringBuilder(length * 2)
			For position As Integer = 0 To length - 2
				Dim current As Char = exampleName.Chars(position)
				Dim [next] As Char = exampleName.Chars(position + 1)
				result.Append(current)
				If Char.IsLower(current) AndAlso Char.IsUpper([next]) Then
					result.Append(" "c)
				End If
			Next position
			result.Append(exampleName.Chars(length - 1))
			Return result.ToString()
		End Function
		Public Shared Function GetExamplePath(ByVal exampleFolderName As String) As String
			Dim examplesPath2 As String = Path.Combine(Directory.GetCurrentDirectory() & "\..\..\", exampleFolderName)
			If Directory.Exists(examplesPath2) Then
				Return examplesPath2
			End If
			Dim examplesPathInInsallation As String = GetRelativeDirectoryPath(exampleFolderName)
			Return examplesPathInInsallation
		End Function
		'public static string GetExamplePath() {
		'    string examplesPath2 = Path.Combine(Directory.GetCurrentDirectory() + "\\..\\..\\", "CodeExamples");
		'    if (Directory.Exists(examplesPath2))
		'        return examplesPath2;
		'    string examplesPathInInsallation = GetRelativeDirectoryPath("CodeExamples");
		'    return examplesPathInInsallation;
		'}
		Public Shared Function GetRelativeDirectoryPath(ByVal name As String) As String
			name = "Data\" & name
			Dim path As String = System.Windows.Forms.Application.StartupPath
			Dim s As String = "\"
			For i As Integer = 0 To 10
				If System.IO.Directory.Exists(path & s & name) Then
					Return (path & s & name)
				Else
					s &= "..\"
				End If
			Next i
			Return ""
		End Function
		Public Shared Function FindExamples(ByVal examplePath As String, ByVal examplesCS As Dictionary(Of String, FileInfo), ByVal examplesVB As Dictionary(Of String, FileInfo)) As List(Of CodeExampleGroup)

			Dim result As New List(Of CodeExampleGroup)()

			Dim current As Dictionary(Of String, FileInfo) = Nothing
			Dim csExampleFinder As ExampleFinder
			Dim vbExampleFinder As ExampleFinder

			If examplesCS.Count = 0 Then
				current = examplesVB
				csExampleFinder = Nothing
				vbExampleFinder = New ExampleFinderVB()
			ElseIf examplesVB.Count = 0 Then
				current = examplesCS
				csExampleFinder = New ExampleFinderCSharp()
				vbExampleFinder = Nothing
			Else
				current = examplesCS
				csExampleFinder = New ExampleFinderCSharp()
				vbExampleFinder = New ExampleFinderVB()
			End If

			For Each sourceCodeItem As KeyValuePair(Of String, FileInfo) In current
				Dim key As String = sourceCodeItem.Key

				Dim findedExamplesCS As New List(Of CodeExample)()
				If csExampleFinder IsNot Nothing Then
					findedExamplesCS = csExampleFinder.Process(examplesCS(key))
				End If

				Dim findedExamplesVB As New List(Of CodeExample)()
				If vbExampleFinder IsNot Nothing Then
					findedExamplesVB = vbExampleFinder.Process(examplesVB(key))
				End If

				Dim mergedExamples As New List(Of CodeExample)()

				If findedExamplesCS.Count <> 0 AndAlso findedExamplesVB.Count = 0 Then
					mergedExamples = findedExamplesCS
				ElseIf findedExamplesCS.Count = 0 AndAlso findedExamplesVB.Count <> 0 Then
					mergedExamples = findedExamplesVB
				ElseIf (findedExamplesCS.Count = findedExamplesVB.Count) Then
					mergedExamples = MergeExamples(findedExamplesCS, findedExamplesVB)
				End If

				If mergedExamples.Count = 0 Then
					Continue For
				End If

				Dim group As New CodeExampleGroup() With {.Name = mergedExamples(0).HumanReadableGroupName, .Examples = mergedExamples}
				result.Add(group)
			Next sourceCodeItem
			Return result
		End Function
		Private Shared Function MergeExamples(ByVal findedExamplesCS As List(Of CodeExample), ByVal findedExamplesVB As List(Of CodeExample)) As List(Of CodeExample)
			Dim result As New List(Of CodeExample)()

			Dim count As Integer = findedExamplesCS.Count
			For i As Integer = 0 To count - 1
				Dim itemCS As CodeExample = findedExamplesCS(i)

				Dim itemVB As CodeExample = findedExamplesVB(i)
				If itemCS.HumanReadableGroupName = itemVB.HumanReadableGroupName AndAlso itemCS.RegionName = itemVB.RegionName Then
					Dim merged As New CodeExample()
					merged.RegionName = itemCS.RegionName
					merged.HumanReadableGroupName = itemCS.HumanReadableGroupName
					merged.CodeCS = itemCS.CodeCS
					merged.CodeVB = itemVB.CodeVB
					result.Add(merged)
				Else
					Throw New InvalidOperationException()
				End If
			Next i
			Return result
		End Function
		Public Shared Function DetectExampleLanguage(ByVal solutionFileNameWithoutExtenstion As String) As ExampleLanguage
			Dim projectPath As String = Directory.GetCurrentDirectory() & "\..\..\"

			Dim csproject() As String = Directory.GetFiles(projectPath, "*.csproj")
			If csproject.Length <> 0 AndAlso csproject(0).EndsWith(solutionFileNameWithoutExtenstion & ".csproj") Then
				Return ExampleLanguage.Csharp
			End If
			Dim vbproject() As String = Directory.GetFiles(projectPath, "*.vbproj")
			If vbproject.Length <> 0 AndAlso vbproject(0).EndsWith(solutionFileNameWithoutExtenstion & ".vbproj") Then
				Return ExampleLanguage.VB
			End If
			Return ExampleLanguage.Csharp
		End Function
	End Class
	#End Region

	#Region "ExampleFinder"
	Public MustInherit Class ExampleFinder
		Public MustOverride ReadOnly Property RegexRegionPattern() As String
		Public MustOverride ReadOnly Property RegionStarts() As String

		Public Function Process(ByVal fileWithExample As FileInfo) As List(Of CodeExample)
			If fileWithExample Is Nothing Then
				Return New List(Of CodeExample)()
			End If

			Dim groupName As String = Path.GetFileNameWithoutExtension(fileWithExample.Name)
			Dim code As String
			Using stream As FileStream = File.Open(fileWithExample.FullName, FileMode.Open, FileAccess.Read)
				Dim sr As New StreamReader(stream)
				code = sr.ReadToEnd()
			End Using
			Dim findedExamples As List(Of CodeExample) = ParseSouceFileAndFindRegionsWithExamples(groupName, code)
			Return findedExamples
		End Function
		' todo: remove example group
		Public Function ParseSouceFileAndFindRegionsWithExamples(ByVal groupName As String, ByVal sourceCode As String) As List(Of CodeExample)
			Dim result As New List(Of CodeExample)()
			Dim examplesByRegion As New Dictionary(Of String, CodeExample)()

			Dim matches = Regex.Matches(sourceCode, RegexRegionPattern, RegexOptions.Singleline)

			For Each match In matches
				Dim lines() As String = match.ToString().Split(New String() { Constants.vbCrLf }, StringSplitOptions.None)

				If lines.Length <= 2 Then
					Continue For
				End If
				'string endRegion = lines[lines.Length - 1];
				Dim isAdditionalModules As Boolean = lines(0).Contains("AdditionalModules")
				Dim isCustomAction As Boolean = lines(0).Contains("CustomActionHandler")

				lines = DeleteLeadingWhiteSpacesFromSourceCode(lines, isAdditionalModules OrElse isCustomAction)

				Dim regionName As String = String.Empty
				Dim regionIsValid As Boolean = ValidateRegionName(lines(0).Replace("AdditionalModules", "").Replace("CustomActionHandler", ""), regionName)
				If (Not regionIsValid) Then
					Continue For
				End If

				Dim exampleCode As String = String.Join(Constants.vbCrLf, lines, 1, lines.Length - 2)
				If isAdditionalModules Then
					If examplesByRegion.ContainsKey(regionName) Then
						SetExampleAdditionalModules(exampleCode, examplesByRegion(regionName))
					End If
				Else
					If isCustomAction Then
						If examplesByRegion.ContainsKey(regionName) Then
							SetExampleCustomActionHandler(exampleCode, examplesByRegion(regionName))
						End If
					Else
						Dim newCodeExample As CodeExample = CreateRichEditExample(groupName, regionName, exampleCode)
						result.Add(newCodeExample)
						examplesByRegion.Add(regionName, newCodeExample)
					End If
				End If
			Next match
			Return result
		End Function

		Protected Function CreateRichEditExample(ByVal exampleGroup As String, ByVal regionName As String, ByVal exampleCode As String) As CodeExample
			Dim result As New CodeExample()
			SetExampleCode(exampleCode, result)
			result.RegionName = regionName
			result.HumanReadableGroupName = CodeExampleDemoUtils.ConvertStringToMoreHumanReadableForm(exampleGroup)
			result.HasCustomAction = False
			Return result
		End Function
		Protected MustOverride Sub SetExampleCode(ByVal exampleCode As String, ByVal newExample As CodeExample)

		Protected MustOverride Sub SetExampleAdditionalModules(ByVal additionalModules As String, ByVal newExample As CodeExample)

		Protected MustOverride Sub SetExampleCustomActionHandler(ByVal customActionHandler As String, ByVal newExample As CodeExample)

		Protected Overridable Function DeleteLeadingWhiteSpacesFromSourceCode(ByVal lines() As String, ByVal isAdditionalModule As Boolean) As String()
            Return If(isAdditionalModule, CodeExampleDemoUtils.DeleteLeadingWhiteSpaces(lines, Constants.vbTab & Constants.vbTab), CodeExampleDemoUtils.DeleteLeadingWhiteSpaces(lines, Constants.vbTab & Constants.vbTab & Constants.vbTab))
		End Function
		Protected Overridable Function ValidateRegionName(ByVal regionLine As String, ByRef regionName As String) As Boolean
			Dim regionIndex As Integer = regionLine.IndexOf(RegionStarts)

			If regionIndex < 0 Then
				regionName = String.Empty
				Return False
			End If

			Dim keepHashMark As Integer = 0 ' "#example" if value is -1 or "example" if value will be 0

			regionName = CodeExampleDemoUtils.ConvertStringToMoreHumanReadableForm(regionLine.Substring(regionIndex + RegionStarts.Length + keepHashMark))
			Return True
		End Function
	End Class
	#End Region
	#Region "ExampleFinderVB"
	Public Class ExampleFinderVB
		Inherits ExampleFinder
		'public ExampleFinderVB() {
		'}
		Public Overrides ReadOnly Property RegexRegionPattern() As String
			Get
				Return "#Region.*?#End Region"
			End Get
		End Property
		Public Overrides ReadOnly Property RegionStarts() As String
			Get
				Return "#Region ""#"
			End Get
		End Property

		Protected Overrides Function DeleteLeadingWhiteSpacesFromSourceCode(ByVal lines() As String, ByVal isAdditionalModule As Boolean) As String()
			Dim result() As String = MyBase.DeleteLeadingWhiteSpacesFromSourceCode(lines, isAdditionalModule)
			Return CodeExampleDemoUtils.DeleteLeadingWhiteSpaces(result, Constants.vbTab + Constants.vbTab + Constants.vbTab)
		End Function
		Protected Overrides Function ValidateRegionName(ByVal regionLine As String, ByRef regionName As String) As Boolean
			Dim result As Boolean = MyBase.ValidateRegionName(regionLine, regionName)
			If (Not result) Then
				Return result
			End If
			regionName = regionName.TrimEnd(""""c)
			Return True
		End Function
		Protected Overrides Sub SetExampleCode(ByVal code As String, ByVal newExample As CodeExample)
			newExample.CodeVB = code
		End Sub

		Protected Overrides Sub SetExampleAdditionalModules(ByVal additionalModules As String, ByVal newExample As CodeExample)
			newExample.AdditionalModulesVB = additionalModules
		End Sub

		Protected Overrides Sub SetExampleCustomActionHandler(ByVal customActionHandler As String, ByVal newExample As CodeExample)
			newExample.HasCustomAction = True
			newExample.CustomActionHandlerVB = customActionHandler
		End Sub
	End Class
	#End Region
	#Region "ExampleFinderCSharp"
	Public Class ExampleFinderCSharp
		Inherits ExampleFinder
		Public Overrides ReadOnly Property RegexRegionPattern() As String
			Get
				Return "#region.*?#endregion"
			End Get
		End Property
		Public Overrides ReadOnly Property RegionStarts() As String
			Get
				Return "#region #"
			End Get
		End Property

		Protected Overrides Sub SetExampleCode(ByVal code As String, ByVal newExample As CodeExample)
			newExample.CodeCS = code
		End Sub

		Protected Overrides Sub SetExampleAdditionalModules(ByVal additionalModules As String, ByVal newExample As CodeExample)
			newExample.AdditionalModulesCS = additionalModules
		End Sub

		Protected Overrides Sub SetExampleCustomActionHandler(ByVal customActionHandler As String, ByVal newExample As CodeExample)
			newExample.HasCustomAction = True
			newExample.CustomActionHandlerCS = customActionHandler
		End Sub

	End Class
	#End Region

	#Region "LeakSafeCompileEventRouter"
	Public Class LeakSafeCompileEventRouter
		Private ReadOnly weakControlRef As WeakReference

		Public Sub New(ByVal [module] As ExampleEvaluatorByTimer)
			'Guard.ArgumentNotNull(module, "module");
			Me.weakControlRef = New WeakReference([module])
		End Sub
		Public Sub OnCompileExampleTimerTick(ByVal sender As Object, ByVal e As EventArgs)
			Dim [module] As ExampleEvaluatorByTimer = CType(weakControlRef.Target, ExampleEvaluatorByTimer)
			If [module] IsNot Nothing Then
				[module].CompileExample(sender, e)
			End If
		End Sub
	End Class
	Public Class CodeEvaluationEventArgs
		Inherits EventArgs
		Private privateResult As Boolean
		Public Property Result() As Boolean
			Get
				Return privateResult
			End Get
			Set(ByVal value As Boolean)
				privateResult = value
			End Set
		End Property
		Private privateCode As String
		Public Property Code() As String
			Get
				Return privateCode
			End Get
			Set(ByVal value As String)
				privateCode = value
			End Set
		End Property
		Private privateAdditionalModules As String
		Public Property AdditionalModules() As String
			Get
				Return privateAdditionalModules
			End Get
			Set(ByVal value As String)
				privateAdditionalModules = value
			End Set
		End Property
		Private privateCustomActionHandler As String
		Public Property CustomActionHandler() As String
			Get
				Return privateCustomActionHandler
			End Get
			Set(ByVal value As String)
				privateCustomActionHandler = value
			End Set
		End Property
		Private privateLanguage As ExampleLanguage
		Public Property Language() As ExampleLanguage
			Get
				Return privateLanguage
			End Get
			Set(ByVal value As ExampleLanguage)
				privateLanguage = value
			End Set
		End Property
		Private privateEvaluationParameter As Object
		Public Property EvaluationParameter() As Object
			Get
				Return privateEvaluationParameter
			End Get
			Set(ByVal value As Object)
				privateEvaluationParameter = value
			End Set
		End Property
		Private privateAdditionalParameter As Object
		Public Property AdditionalParameter() As Object
			Get
				Return privateAdditionalParameter
			End Get
			Set(ByVal value As Object)
				privateAdditionalParameter = value
			End Set
		End Property
	End Class
	Public Delegate Sub CodeEvaluationEventHandler(ByVal sender As Object, ByVal e As CodeEvaluationEventArgs)

	Public Class OnAfterCompileEventArgs
		Inherits EventArgs
		Private privateResult As Boolean
		Public Property Result() As Boolean
			Get
				Return privateResult
			End Get
			Set(ByVal value As Boolean)
				privateResult = value
			End Set
		End Property
	End Class
	Public Delegate Sub OnAfterCompileEventHandler(ByVal sender As Object, ByVal e As OnAfterCompileEventArgs)
	#End Region

	Public MustInherit Class ExampleEvaluatorByTimer
		Implements IDisposable
		Private leakSafeCompileEventRouter As LeakSafeCompileEventRouter
		Private compileExampleTimer As System.Windows.Forms.Timer
		Private compileComplete As Boolean = True
		Private Const CompileTimeIntervalInMilliseconds As Integer = 2000

		Public Sub New(ByVal enableTimer As Boolean)
			Me.leakSafeCompileEventRouter = New LeakSafeCompileEventRouter(Me)

			'this.compileExampleTimer = new System.Windows.Forms.Timer();
			If enableTimer Then
				Me.compileExampleTimer = New System.Windows.Forms.Timer()
				Me.compileExampleTimer.Interval = CompileTimeIntervalInMilliseconds

				'this.compileExampleTimer.Tick += new EventHandler(leakSafeCompileEventRouter.OnCompileExampleTimerTick); //OnCompileTimerTick
				Me.compileExampleTimer.Enabled = True
			End If
		End Sub
		Public Sub New()
			Me.New(True)
		End Sub

		#Region "Events"
		Public Event QueryEvaluate As CodeEvaluationEventHandler
		'public event CodeEvaluationEventHandler QueryEvaluateEvent {
		'    add { onQeuryEvaluate += value; }
		'    remove { onQeuryEvaluate -= value; }
		'}
		Protected Friend Overridable Function RaiseQueryEvaluate() As CodeEvaluationEventArgs
			If QueryEvaluateEvent IsNot Nothing Then
				Dim args As New CodeEvaluationEventArgs()
				RaiseEvent QueryEvaluate(Me, args)
				Return args
			End If
			Return Nothing
		End Function
		Public Event OnBeforeCompile As EventHandler
		'public event EventHandler OnBeforeCompileEvent { add { onBeforeCompile += value; } remove { onBeforeCompile -= value; } }
		Private Sub RaiseOnBeforeCompile()
			RaiseEvent OnBeforeCompile(Me, New EventArgs())
		End Sub

		Public Event OnAfterCompile As OnAfterCompileEventHandler
		'public event OnAfterCompileEventHandler OnAfterCompileEvent { add { onAfterCompile += value; } remove { onAfterCompile -= value; } }
		Private Sub RaiseOnAfterCompile(ByVal result As Boolean)
			RaiseEvent OnAfterCompile(Me, New OnAfterCompileEventArgs() With {.Result = result})
		End Sub
		#End Region

		Public Sub CompileExample(ByVal sender As Object, ByVal e As EventArgs)
			If (Not compileComplete) Then
				Return
			End If
			Dim args As CodeEvaluationEventArgs = RaiseQueryEvaluate()
			If (Not args.Result) Then
				Return
			End If

			ForceCompile(args)
		End Sub
		Public Sub ForceCompile(ByVal args As CodeEvaluationEventArgs)
			compileComplete = False
			If (Not String.IsNullOrEmpty(args.Code)) Then
				CompileExampleAndShowPrintPreview(args)
			End If

			compileComplete = True
		End Sub
		Private Sub CompileExampleAndShowPrintPreview(ByVal args As CodeEvaluationEventArgs)
			Dim evaluationSucceed As Boolean = False
			Try
				RaiseOnBeforeCompile()

				evaluationSucceed = Evaluate(args)
			Finally
				RaiseOnAfterCompile(evaluationSucceed)
			End Try
		End Sub

		Public Function Evaluate(ByVal args As CodeEvaluationEventArgs) As Boolean
			Dim richeditExampleCodeEvaluator As ExampleCodeEvaluator = GetExampleCodeEvaluator(args.Language)
			Return richeditExampleCodeEvaluator.ExcecuteCodeAndGenerateDocument(args)
		End Function

		Protected MustOverride Function GetExampleCodeEvaluator(ByVal language As ExampleLanguage) As ExampleCodeEvaluator

		Public Sub Dispose() Implements IDisposable.Dispose
			If compileExampleTimer IsNot Nothing Then
				compileExampleTimer.Enabled = False
				If leakSafeCompileEventRouter IsNot Nothing Then
					RemoveHandler compileExampleTimer.Tick, AddressOf leakSafeCompileEventRouter.OnCompileExampleTimerTick 'OnCompileTimerTick
				End If
				compileExampleTimer.Dispose()
				compileExampleTimer = Nothing
			End If
		End Sub
	End Class
End Namespace
