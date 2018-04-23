using System.CodeDom.Compiler;
using DevExpress.Internal;
using System.Reflection;
using System.IO;
using System.Text.RegularExpressions;
using System;
using System.Collections.Generic;
using System.Text;

namespace RichEditAPISample
{
    public abstract class ExampleCodeEvaluator
    {
        protected abstract string CodeStart { get; }
        protected abstract string MainMethodEnd { get; }
        protected abstract string CodeEnd { get; }
        protected abstract CodeDomProvider GetCodeDomProvider();
        protected abstract string[] GetModuleAssemblies();
        protected abstract string GetExampleClassName();

        public bool ExcecuteCodeAndGenerateDocument(CodeEvaluationEventArgs args)
        {
            string theCode = String.Concat(CodeStart, args.Code, MainMethodEnd, args.CustomActionHandler, args.AdditionalModules, CodeEnd);
            string[] linesOfCode = new string[] { theCode };
            return CompileAndRun(linesOfCode, args.EvaluationParameter, args.AdditionalParameter);
        }

        protected internal bool CompileAndRun(string[] linesOfCode, object evaluationParameter, object additionalParameter)
        {
            CompilerParameters CompilerParams = new CompilerParameters();

            CompilerParams.GenerateInMemory = true;
            CompilerParams.TreatWarningsAsErrors = false;
            CompilerParams.GenerateExecutable = false;

            string[] referencesSystem = new string[] { "System.dll",
                                                      "System.Windows.Forms.dll",
                                                      "System.Data.dll",
                                                      "System.Xml.dll",
                                                      "System.Core.dll",
                                                      "System.Drawing.dll" };

            string[] referencesDX = new string[] {
                AssemblyInfo.SRAssemblyData,
                AssemblyInfo.SRAssemblyOfficeCore,
                AssemblyInfo.SRAssemblyPrintingCore,
                AssemblyInfo.SRAssemblyPrinting,
                AssemblyInfo.SRAssemblyDocs,
                AssemblyInfo.SRAssemblyUtils
            };

            string[] referencesDXModule = GetModuleAssemblies();

            string[] references = new string[referencesSystem.Length + referencesDX.Length + referencesDXModule.Length];

            for (int referenceIndex = 0; referenceIndex < referencesSystem.Length; referenceIndex++)
            {
                references[referenceIndex] = referencesSystem[referenceIndex];
            }

            for (int i = 0, initial = referencesSystem.Length; i < referencesDX.Length; i++)
            {
                Assembly assembly = Assembly.Load(referencesDX[i] + AssemblyInfo.FullAssemblyVersionExtension);
                if (assembly != null)
                    references[i + initial] = assembly.Location;
            }

            for(int i = 0, initial = referencesSystem.Length + referencesDX.Length; i < referencesDXModule.Length; i++) {
                Assembly assembly = Assembly.Load(referencesDXModule[i] + AssemblyInfo.FullAssemblyVersionExtension);
                if(assembly != null)
                    references[i + initial] = assembly.Location;
            }

            CompilerParams.ReferencedAssemblies.AddRange(references);


            CodeDomProvider provider = GetCodeDomProvider();
            CompilerResults compile = provider.CompileAssemblyFromSource(CompilerParams, linesOfCode);

            if (compile.Errors.HasErrors)
            {
                //string text = "Compile error: ";
                //foreach(CompilerError ce in compile.Errors) {
                //    text += "rn" + ce.ToString();
                //}
                //MessageBox.Show(text);
                return false;
            }

            Module module = null;
            try
            {
                module = compile.CompiledAssembly.GetModules()[0];
            }
            catch
            {
            }
            Type moduleType = null;
            if (module == null)
            {
                return false;
            }
            moduleType = module.GetType(GetExampleClassName());

            MethodInfo methInfo = null;
            if (moduleType == null)
            {
                return false;
            }
            methInfo = moduleType.GetMethod("Process");

            if (methInfo != null)
            {
                try
                {
                    methInfo.Invoke(null, new object[] { evaluationParameter, additionalParameter });
                }
                catch (Exception)
                {
                    return false;// an error
                }
                return true;
            }
            return false;
        }
    }

    public class CodeExampleGroup
    {
        public CodeExampleGroup()
        {
        }
        public string Name { get; set; }
        public List<CodeExample> Examples { get; set; }
        public int Id { get; set; }
    }

    public class CodeExample
    {
        public string CodeCS { get; set; }
        public string CodeVB { get; set; }
        public string AdditionalModulesCS { get; set; }
        public string AdditionalModulesVB { get; set; }
        public string RegionName { get; set; }
        public string HumanReadableGroupName { get; set; }
        public string ExampleGroup { get; set; }
        public bool HasCustomAction { get; set; }
        public string CustomActionHandlerCS { get; set; }
        public string CustomActionHandlerVB { get; set; }
        public int Id { get; set; }
    }

    public enum ExampleLanguage
    {
        Csharp = 0,
        VB = 1
    }

    #region CodeExampleDemoUtils
    public static class CodeExampleDemoUtils
    {
        public static Dictionary<string, FileInfo> GatherExamplesFromProject(string examplesPath, ExampleLanguage language)
        {
            Dictionary<string, FileInfo> result = new Dictionary<string, FileInfo>();
            foreach (string fileName in Directory.GetFiles(examplesPath, "*" + GetCodeExampleFileExtension(language)))
                result.Add(Path.GetFileNameWithoutExtension(fileName), new FileInfo(fileName));
            return result;
        }
        public static string GetCodeExampleFileExtension(ExampleLanguage language)
        {
            if (language == ExampleLanguage.VB)
                return ".vb";
            return ".cs";
        }
        public static string[] DeleteLeadingWhiteSpaces(string[] lines, String stringToDelete)
        {
            string[] result = new string[lines.Length];
            int stringToDeleteLength = stringToDelete.Length;

            for (int i = 0; i < lines.Length; i++)
            {
                int index = lines[i].IndexOf(stringToDelete);
                result[i] = (index >= 0) ? lines[i].Substring(index + stringToDeleteLength) : lines[i];
            }
            return result;
        }
        public static string ConvertStringToMoreHumanReadableForm(string exampleName)
        {
            string result = SplitCamelCase(exampleName);
            result = result.Replace(" In ", " in ");
            result = result.Replace(" And ", " and ");
            result = result.Replace(" To ", " to ");
            result = result.Replace(" From ", " from ");
            result = result.Replace(" With ", " with ");
            result = result.Replace(" By ", " by ");
            return result;
        }
        static string SplitCamelCase(string exampleName)
        {
            int length = exampleName.Length;
            if (length == 1)
                return exampleName;

            StringBuilder result = new StringBuilder(length * 2);
            for (int position = 0; position < length - 1; position++)
            {
                char current = exampleName[position];
                char next = exampleName[position + 1];
                result.Append(current);
                if (char.IsLower(current) && char.IsUpper(next))
                {
                    result.Append(' ');
                }
            }
            result.Append(exampleName[length - 1]);
            return result.ToString();
        }
        public static string GetExamplePath(string exampleFolderName)
        {//"CodeExamples"
            string examplesPath2 = Path.Combine(Directory.GetCurrentDirectory() + "\\..\\..\\", exampleFolderName);
            if (Directory.Exists(examplesPath2))
                return examplesPath2;
            string examplesPathInInsallation = GetRelativeDirectoryPath(exampleFolderName);
            return examplesPathInInsallation;
        }
        //public static string GetExamplePath() {
        //    string examplesPath2 = Path.Combine(Directory.GetCurrentDirectory() + "\\..\\..\\", "CodeExamples");
        //    if (Directory.Exists(examplesPath2))
        //        return examplesPath2;
        //    string examplesPathInInsallation = GetRelativeDirectoryPath("CodeExamples");
        //    return examplesPathInInsallation;
        //}
        public static string GetRelativeDirectoryPath(string name)
        {
            name = "Data\\" + name;
            string path = System.Windows.Forms.Application.StartupPath;
            string s = "\\";
            for (int i = 0; i <= 10; i++)
            {
                if (System.IO.Directory.Exists(path + s + name))
                    return (path + s + name);
                else
                    s += "..\\";
            }
            return "";
        }
        public static List<CodeExampleGroup> FindExamples(string examplePath, Dictionary<string, FileInfo> examplesCS, Dictionary<string, FileInfo> examplesVB)
        {

            List<CodeExampleGroup> result = new List<CodeExampleGroup>();

            Dictionary<string, FileInfo> current = null;
            ExampleFinder csExampleFinder;
            ExampleFinder vbExampleFinder;

            if (examplesCS.Count == 0)
            {
                current = examplesVB;
                csExampleFinder = null;
                vbExampleFinder = new ExampleFinderVB();
            }
            else if (examplesVB.Count == 0)
            {
                current = examplesCS;
                csExampleFinder = new ExampleFinderCSharp();
                vbExampleFinder = null;
            }
            else
            {
                current = examplesCS;
                csExampleFinder = new ExampleFinderCSharp();
                vbExampleFinder = new ExampleFinderVB();
            }

            foreach (KeyValuePair<string, FileInfo> sourceCodeItem in current)
            {
                string key = sourceCodeItem.Key;

                List<CodeExample> findedExamplesCS = new List<CodeExample>();
                if (csExampleFinder != null)
                    findedExamplesCS = csExampleFinder.Process(examplesCS[key]);

                List<CodeExample> findedExamplesVB = new List<CodeExample>();
                if (vbExampleFinder != null)
                    findedExamplesVB = vbExampleFinder.Process(examplesVB[key]);

                List<CodeExample> mergedExamples = new List<CodeExample>();

                if (findedExamplesCS.Count != 0 && findedExamplesVB.Count == 0)
                    mergedExamples = findedExamplesCS;
                else if (findedExamplesCS.Count == 0 && findedExamplesVB.Count != 0)
                    mergedExamples = findedExamplesVB;
                else if ((findedExamplesCS.Count == findedExamplesVB.Count))
                    mergedExamples = MergeExamples(findedExamplesCS, findedExamplesVB);

                if (mergedExamples.Count == 0)
                    continue;

                CodeExampleGroup group = new CodeExampleGroup()
                {
                    Name = mergedExamples[0].HumanReadableGroupName,
                    Examples = mergedExamples
                };
                result.Add(group);
            }
            return result;
        }
        static List<CodeExample> MergeExamples(List<CodeExample> findedExamplesCS, List<CodeExample> findedExamplesVB)
        {
            List<CodeExample> result = new List<CodeExample>();

            int count = findedExamplesCS.Count;
            for (int i = 0; i < count; i++)
            {
                CodeExample itemCS = findedExamplesCS[i];

                CodeExample itemVB = findedExamplesVB[i];
                if (itemCS.HumanReadableGroupName == itemVB.HumanReadableGroupName
                    && itemCS.RegionName == itemVB.RegionName)
                {
                    CodeExample merged = new CodeExample();
                    merged.RegionName = itemCS.RegionName;
                    merged.HumanReadableGroupName = itemCS.HumanReadableGroupName;
                    merged.CodeCS = itemCS.CodeCS;
                    merged.CodeVB = itemVB.CodeVB;
                    result.Add(merged);
                }
                else
                    throw new InvalidOperationException();
            }
            return result;
        }
        public static ExampleLanguage DetectExampleLanguage(string solutionFileNameWithoutExtenstion)
        {
            string projectPath = Directory.GetCurrentDirectory() + "\\..\\..\\";

            string[] csproject = Directory.GetFiles(projectPath, "*.csproj");
            if (csproject.Length != 0 && csproject[0].EndsWith(solutionFileNameWithoutExtenstion + ".csproj"))
                return ExampleLanguage.Csharp;
            string[] vbproject = Directory.GetFiles(projectPath, "*.vbproj");
            if (vbproject.Length != 0 && vbproject[0].EndsWith(solutionFileNameWithoutExtenstion + ".vbproj"))
                return ExampleLanguage.VB;
            return ExampleLanguage.Csharp;
        }
    }
    #endregion

    #region ExampleFinder
    public abstract class ExampleFinder
    {
        public abstract string RegexRegionPattern { get; }
        public abstract string RegionStarts { get; }

        public List<CodeExample> Process(FileInfo fileWithExample)
        {
            if (fileWithExample == null)
                return new List<CodeExample>();

            string groupName = Path.GetFileNameWithoutExtension(fileWithExample.Name);
            string code;
            using (FileStream stream = File.Open(fileWithExample.FullName, FileMode.Open, FileAccess.Read))
            {
                StreamReader sr = new StreamReader(stream);
                code = sr.ReadToEnd();
            }
            List<CodeExample> findedExamples = ParseSouceFileAndFindRegionsWithExamples(groupName, code);
            return findedExamples;
        }
        // todo: remove example group
        public List<CodeExample> ParseSouceFileAndFindRegionsWithExamples(string groupName, string sourceCode)
        {
            List<CodeExample> result = new List<CodeExample>();
            Dictionary<string, CodeExample> examplesByRegion = new Dictionary<string, CodeExample>();

            var matches = Regex.Matches(sourceCode, RegexRegionPattern, RegexOptions.Singleline);

            foreach (var match in matches)
            {
                string[] lines = match.ToString().Split(new string[] { "\r\n" }, StringSplitOptions.None);

                if (lines.Length <= 2)
                    continue;
                //string endRegion = lines[lines.Length - 1];
                bool isAdditionalModules = lines[0].Contains("AdditionalModules");
                bool isCustomAction = lines[0].Contains("CustomActionHandler");

                lines = DeleteLeadingWhiteSpacesFromSourceCode(lines, isAdditionalModules || isCustomAction);

                string regionName = String.Empty;
                bool regionIsValid = ValidateRegionName(lines[0].Replace("AdditionalModules", "").Replace("CustomActionHandler", ""), ref regionName);
                if (!regionIsValid)
                    continue;

                string exampleCode = string.Join("\r\n", lines, 1, lines.Length - 2);
                if(isAdditionalModules) {
                    if(examplesByRegion.ContainsKey(regionName)) SetExampleAdditionalModules(exampleCode, examplesByRegion[regionName]);
                }
                else {
                    if(isCustomAction) {
                        if(examplesByRegion.ContainsKey(regionName)) SetExampleCustomActionHandler(exampleCode, examplesByRegion[regionName]);
                    }
                    else {
                        CodeExample newCodeExample = CreateRichEditExample(groupName, regionName, exampleCode);
                        result.Add(newCodeExample);
                        examplesByRegion.Add(regionName, newCodeExample);
                    }
                }
            }
            return result;
        }

        protected CodeExample CreateRichEditExample(string exampleGroup, string regionName, string exampleCode)
        {
            CodeExample result = new CodeExample();
            SetExampleCode(exampleCode, result);
            result.RegionName = regionName;
            result.HumanReadableGroupName = CodeExampleDemoUtils.ConvertStringToMoreHumanReadableForm(exampleGroup);
            result.HasCustomAction = false;
            return result;
        }
        protected abstract void SetExampleCode(string exampleCode, CodeExample newExample);

        protected abstract void SetExampleAdditionalModules(string additionalModules, CodeExample newExample);

        protected abstract void SetExampleCustomActionHandler(string customActionHandler, CodeExample newExample);

        protected virtual string[] DeleteLeadingWhiteSpacesFromSourceCode(string[] lines, bool isAdditionalModule)
        {
            return isAdditionalModule ? CodeExampleDemoUtils.DeleteLeadingWhiteSpaces(lines, "        ") : CodeExampleDemoUtils.DeleteLeadingWhiteSpaces(lines, "            ");
        }
        protected virtual bool ValidateRegionName(string regionLine, ref string regionName)
        {
            int regionIndex = regionLine.IndexOf(RegionStarts);

            if (regionIndex < 0)
            {
                regionName = String.Empty;
                return false;
            }

            int keepHashMark = 0; // "#example" if value is -1 or "example" if value will be 0

            regionName = CodeExampleDemoUtils.ConvertStringToMoreHumanReadableForm(regionLine.Substring(regionIndex + RegionStarts.Length + keepHashMark));
            return true;
        }
    }
    #endregion
    #region ExampleFinderVB
    public class ExampleFinderVB : ExampleFinder
    {
        //public ExampleFinderVB() {
        //}
        public override string RegexRegionPattern { get { return "#Region.*?#End Region"; } }
        public override string RegionStarts { get { return "#Region \"#"; } }

        protected override string[] DeleteLeadingWhiteSpacesFromSourceCode(string[] lines, bool isAdditionalModule)
        {
            string[] result = base.DeleteLeadingWhiteSpacesFromSourceCode(lines, isAdditionalModule);
            return CodeExampleDemoUtils.DeleteLeadingWhiteSpaces(result, "\t\t\t");
        }
        protected override bool ValidateRegionName(string regionLine, ref string regionName)
        {
            bool result = base.ValidateRegionName(regionLine, ref regionName);
            if (!result)
                return result;
            regionName = regionName.TrimEnd('\"');
            return true;
        }
        protected override void SetExampleCode(string code, CodeExample newExample)
        {
            newExample.CodeVB = code;
        }

        protected override void SetExampleAdditionalModules(string additionalModules, CodeExample newExample) {
            newExample.AdditionalModulesVB = additionalModules;
        }

        protected override void SetExampleCustomActionHandler(string customActionHandler, CodeExample newExample) {
            newExample.HasCustomAction = true;
            newExample.CustomActionHandlerVB = customActionHandler;
        }
    }
    #endregion
    #region ExampleFinderCSharp
    public class ExampleFinderCSharp : ExampleFinder
    {
        public override string RegexRegionPattern { get { return "#region.*?#endregion"; } }
        public override string RegionStarts { get { return "#region #"; } }

        protected override void SetExampleCode(string code, CodeExample newExample)
        {
            newExample.CodeCS = code;
        }

        protected override void SetExampleAdditionalModules(string additionalModules, CodeExample newExample) {
            newExample.AdditionalModulesCS = additionalModules;
        }

        protected override void SetExampleCustomActionHandler(string customActionHandler, CodeExample newExample) {
            newExample.HasCustomAction = true;
            newExample.CustomActionHandlerCS = customActionHandler;
        }

    }
    #endregion

    #region LeakSafeCompileEventRouter
    public class LeakSafeCompileEventRouter
    {
        readonly WeakReference weakControlRef;

        public LeakSafeCompileEventRouter(ExampleEvaluatorByTimer module)
        {
            //Guard.ArgumentNotNull(module, "module");
            this.weakControlRef = new WeakReference(module);
        }
        public void OnCompileExampleTimerTick(object sender, EventArgs e)
        {
            ExampleEvaluatorByTimer module = (ExampleEvaluatorByTimer)weakControlRef.Target;
            if (module != null)
                module.CompileExample(sender, e);
        }
    }
    public class CodeEvaluationEventArgs : EventArgs
    {
        public bool Result { get; set; }
        public string Code { get; set; }
        public string AdditionalModules { get; set; }
        public string CustomActionHandler { get; set; }
        public ExampleLanguage Language { get; set; }
        public object EvaluationParameter { get; set; }
        public object AdditionalParameter { get; set; }
    }
    public delegate void CodeEvaluationEventHandler(object sender, CodeEvaluationEventArgs e);

    public class OnAfterCompileEventArgs : EventArgs
    {
        public bool Result { get; set; }
    }
    public delegate void OnAfterCompileEventHandler(object sender, OnAfterCompileEventArgs e);
    #endregion

    public abstract class ExampleEvaluatorByTimer : IDisposable
    {
        LeakSafeCompileEventRouter leakSafeCompileEventRouter;
        System.Windows.Forms.Timer compileExampleTimer;
        bool compileComplete = true;
        const int CompileTimeIntervalInMilliseconds = 2000;

        public ExampleEvaluatorByTimer(bool enableTimer)
        {
            this.leakSafeCompileEventRouter = new LeakSafeCompileEventRouter(this);

            //this.compileExampleTimer = new System.Windows.Forms.Timer();
            if (enableTimer)
            {
                this.compileExampleTimer = new System.Windows.Forms.Timer();
                this.compileExampleTimer.Interval = CompileTimeIntervalInMilliseconds;

                //this.compileExampleTimer.Tick += new EventHandler(leakSafeCompileEventRouter.OnCompileExampleTimerTick); //OnCompileTimerTick
                this.compileExampleTimer.Enabled = true;
            }
        }
        public ExampleEvaluatorByTimer()
            : this(true)
        {
        }

        #region Events
        public event CodeEvaluationEventHandler QueryEvaluate;
        //public event CodeEvaluationEventHandler QueryEvaluateEvent {
        //    add { onQeuryEvaluate += value; }
        //    remove { onQeuryEvaluate -= value; }
        //}
        protected internal virtual CodeEvaluationEventArgs RaiseQueryEvaluate()
        {
            if (QueryEvaluate != null)
            {
                CodeEvaluationEventArgs args = new CodeEvaluationEventArgs();
                QueryEvaluate(this, args);
                return args;
            }
            return null;
        }
        public event EventHandler OnBeforeCompile;
        //public event EventHandler OnBeforeCompileEvent { add { onBeforeCompile += value; } remove { onBeforeCompile -= value; } }
        void RaiseOnBeforeCompile()
        {
            if (OnBeforeCompile != null)
                OnBeforeCompile(this, new EventArgs());
        }

        public event OnAfterCompileEventHandler OnAfterCompile;
        //public event OnAfterCompileEventHandler OnAfterCompileEvent { add { onAfterCompile += value; } remove { onAfterCompile -= value; } }
        void RaiseOnAfterCompile(bool result)
        {
            if (OnAfterCompile != null)
                OnAfterCompile(this, new OnAfterCompileEventArgs() { Result = result });
        }
        #endregion

        public void CompileExample(object sender, EventArgs e)
        {
            if (!compileComplete)
                return;
            CodeEvaluationEventArgs args = RaiseQueryEvaluate();
            if (!args.Result)
                return;

            ForceCompile(args);
        }
        public void ForceCompile(CodeEvaluationEventArgs args)
        {
            compileComplete = false;
            if (!String.IsNullOrEmpty(args.Code))
                CompileExampleAndShowPrintPreview(args);

            compileComplete = true;
        }
        void CompileExampleAndShowPrintPreview(CodeEvaluationEventArgs args)
        {
            bool evaluationSucceed = false;
            try
            {
                RaiseOnBeforeCompile();

                evaluationSucceed = Evaluate(args);
            }
            finally
            {
                RaiseOnAfterCompile(evaluationSucceed);
            }
        }

        public bool Evaluate(CodeEvaluationEventArgs args)
        {
            ExampleCodeEvaluator richeditExampleCodeEvaluator = GetExampleCodeEvaluator(args.Language);
            return richeditExampleCodeEvaluator.ExcecuteCodeAndGenerateDocument(args);
        }

        protected abstract ExampleCodeEvaluator GetExampleCodeEvaluator(ExampleLanguage language);

        public void Dispose()
        {
            if (compileExampleTimer != null)
            {
                compileExampleTimer.Enabled = false;
                if (leakSafeCompileEventRouter != null)
                    compileExampleTimer.Tick -= new EventHandler(leakSafeCompileEventRouter.OnCompileExampleTimerTick); //OnCompileTimerTick
                compileExampleTimer.Dispose();
                compileExampleTimer = null;
            }
        }
    }
}
