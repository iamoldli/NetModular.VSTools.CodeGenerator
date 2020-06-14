using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace NetModular.VSTools.CodeGenerator
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class NetModularCommand
    {
        private static DTE2 _dte;

        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("3868efcd-8bcd-4924-aa18-09e1b13de758");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="NetModularCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private NetModularCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static NetModularCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package, DTE2 dte)
        {
            // Switch to the main thread - the call to AddCommand in NetModularCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new NetModularCommand(package, commandService);

            _dte = dte;
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            string message = "";

            try
            {
                if (_dte.SelectedItems.Count > 0)
                {
                    SelectedItem selectedItem = _dte.SelectedItems.Item(1);
                    ProjectItem selectProjectItem = selectedItem.ProjectItem;

                    if (selectProjectItem != null)
                    {
                        #region 获取出基础信息
                        //获取当前点击的类所在的项目
                        Project topProject = selectProjectItem.ContainingProject;

                        //当前类在当前项目中的目录结构
                        string dirPath = GetSelectFileDirPath(topProject, selectProjectItem);

                        //当前类命名空间
                        string namespaceStr = selectProjectItem.FileCodeModel.CodeElements.OfType<CodeNamespace>().First().FullName;

                        //当前项目根命名空间
                        string applicationStr = "";
                        if (!string.IsNullOrEmpty(namespaceStr))
                        {
                            applicationStr = namespaceStr.Substring(0, namespaceStr.IndexOf("."));
                        }

                        //当前类
                        CodeClass codeClass = GetClass(selectProjectItem.FileCodeModel.CodeElements);

                        //当前项目类名
                        string className = codeClass.Name;

                        if (!codeClass.Name.Contains("Repository"))
                            throw new Exception("请在Repository类上操作");

                        //当前类中文名 [Display(Name = "供应商")]
                        string classCnName = "";

                        //当前类说明 [Description("品牌信息")]
                        string classDescription = "";

                        //获取类的中文名称和说明
                        foreach (CodeAttribute classAttribute in codeClass.Attributes)
                        {
                            switch (classAttribute.Name)
                            {
                                case "Display":
                                    if (!string.IsNullOrEmpty(classAttribute.Value))
                                    {
                                        string displayStr = classAttribute.Value.Trim();
                                        foreach (var displayValueStr in displayStr.Split(','))
                                        {
                                            if (!string.IsNullOrEmpty(displayValueStr))
                                            {
                                                if (displayValueStr.Split('=')[0].Trim() == "Name")
                                                {
                                                    classCnName = displayValueStr.Split('=')[1].Trim().Replace("\"", "");
                                                }
                                            }
                                        }
                                    }
                                    break;
                                case "Description":
                                    classDescription = classAttribute.Value;
                                    break;
                            }
                        }

                        #endregion

                        List<ClassFunction> classFuns = GetClassModel(applicationStr, className, classCnName, classDescription, dirPath, codeClass).ClassFunctions;


                        ////获取仓储类Repository中的函数列表
                        //CodeFunction codeFunction = GetFunction(selectProjectItem.FileCodeModel.CodeElements);


                        //去掉类名Repository
                        className = className.Replace("Repository", "");

                        //获取当前解决方案的项目列表
                        List<Project> solutionProjectItems = GetSolutionProjects(_dte.Solution);


                        //获取Application项目对象
                        Project applicationProjectItem = solutionProjectItems.Find(t => t.FullName == topProject.FullName.Replace("Infrastructure", "Application"));
                        SetIService(applicationProjectItem, classFuns, className);

                        SetService(applicationProjectItem, classFuns, className);

                        //获取Domain项目对象
                        Project domainProjectItem = solutionProjectItems.Find(t => t.FullName == topProject.FullName.Replace("Infrastructure", "Domain"));
                        SetIRepository(domainProjectItem, classFuns, className);


                        message = "生成成功！";
                    }
                }

            }
            catch (Exception ex)
            {
                message = ex.Message;
            }


            string title = "NetModular框架代码生成器";

            // Show a message box to prove we were here
            VsShellUtilities.ShowMessageBox(
                this.package,
                message,
                title,
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

        }


        #region 代码生成

        /// <summary>
        /// 生成Domain项目IRepository类方法
        /// </summary>
        /// <param name="topProject"></param>
        /// <param name="className"></param>
        private void SetIRepository(Project applicationProject, List<ClassFunction> classFuns, string className)
        {
            var classFileName = applicationProject.FileName.Substring(0, applicationProject.FileName.LastIndexOf("\\")) + "\\" + className + "\\I" + className + "Repository.cs";

            ProjectItem projectItem = _dte.Solution.FindProjectItem(classFileName);
            if (projectItem != null)
            {
                var codeInterface = GetInterface(projectItem.FileCodeModel.CodeElements);
                var codeChilds = codeInterface?.Collection;
                if (codeChilds == null)
                    return;

                //取第一个类
                CodeElement codeChild = codeChilds.Cast<CodeElement>().First();
                var insertCode = codeChild.GetEndPoint(vsCMPart.vsCMPartBody).CreateEditPoint();

                //获取待修改类的函数结构
                List<ClassFunction> funs = GetInterfaceModel("", "", "", "", "", codeInterface).ClassFunctions;

                foreach (ClassFunction fun in classFuns)
                {
                    //如果函数名称和参数一致,则不用生成代码
                    if (funs.Exists(m => m.Name == fun.Name && m.ClassParametersCast == fun.ClassParametersCast))
                        continue;

                    insertCode.Insert("\r\n");
                    insertCode.Insert("        " + fun.FunctionType + " " + fun.Name + "(" + fun.ClassParametersCast + ");\r\n");
                    insertCode.Insert("\r\n");
                }

                projectItem.Save();
            }
        }

        /// <summary>
        /// 生成Application项目IService类方法
        /// </summary>
        /// <param name="topProject"></param>
        /// <param name="className"></param>
        private void SetIService(Project applicationProject, List<ClassFunction> classFuns, string className)
        {
            //类名加上约定的Service后缀
            className += "Service";
            var classFileName = applicationProject.FileName.Substring(0, applicationProject.FileName.LastIndexOf("\\")) + "\\" + className + "\\I" + className + ".cs";

            ProjectItem projectItem = _dte.Solution.FindProjectItem(classFileName);
            if (projectItem != null)
            {
                var codeInterface = GetInterface(projectItem.FileCodeModel.CodeElements);
                var codeChilds = codeInterface?.Collection;
                if (codeChilds == null)
                    return;

                //取第一个类
                CodeElement codeChild = codeChilds.Cast<CodeElement>().First();
                var insertCode = codeChild.GetEndPoint(vsCMPart.vsCMPartBody).CreateEditPoint();

                //获取待修改类的函数结构
                List<ClassFunction> funs = GetInterfaceModel("", "", "", "", "", codeInterface).ClassFunctions;

                foreach (ClassFunction fun in classFuns)
                {
                    //如果函数名称和参数一致,则不用生成代码
                    if (funs.Exists(m => m.Name == fun.Name && m.ClassParametersCast == fun.ClassParametersCast))
                        continue;


                    //返回实际数据对象
                    insertCode.Insert("\r\n");
                    insertCode.Insert("        " + fun.FunctionType + " " + fun.Name + "(" + Cast(fun.ClassParameters) + ");\r\n");
                    insertCode.Insert("\r\n");


                    //返回IResultModel对象
                    //insertCode.Insert("\r\n");
                    //insertCode.Insert("        Task<IResultModel> " + fun.Name + "(" + fun.ClassParametersCast + ");\r\n");
                    //insertCode.Insert("\r\n");
                }

                projectItem.Save();
            }
        }

        /// <summary>
        /// 生成Application项目Service类方法
        /// </summary>
        /// <param name="topProject"></param>
        /// <param name="className"></param>
        private void SetService(Project applicationProject, List<ClassFunction> classFuns, string className)
        {
            //类名加上约定的Service后缀
            className += "Service";
            var classFileName = applicationProject.FileName.Substring(0, applicationProject.FileName.LastIndexOf("\\")) + "\\" + className + "\\" + className + ".cs";

            ProjectItem projectItem = _dte.Solution.FindProjectItem(classFileName);
            if (projectItem != null)
            {
                var codeClass = GetClass(projectItem.FileCodeModel.CodeElements);
                var codeChilds = codeClass?.Collection;
                if (codeChilds == null)
                    return;

                //取第一个类
                CodeElement codeChild = codeChilds.Cast<CodeElement>().First();
                var insertCode = codeChild.GetEndPoint(vsCMPart.vsCMPartBody).CreateEditPoint();

                //获取待修改类的函数结构
                List<ClassFunction> funs = GetClassModel("", "", "", "", "", codeClass).ClassFunctions;

                foreach (ClassFunction fun in classFuns)
                {
                    //如果函数名称和参数一致,则不用生成代码
                    if (funs.Exists(m => m.Name == fun.Name && m.ClassParametersCast == fun.ClassParametersCast))
                        continue;

                    //返回实际数据对象
                    insertCode.Insert("\r\n");
                    insertCode.Insert("        " + fun.Access + " " + (fun.FunctionType.Contains("Task") ? "async" : "") + " " + fun.FunctionType + " " + fun.Name + "(" + fun.ClassParametersCast + ")\r\n");
                    insertCode.Insert("        {\r\n");
                    insertCode.Insert("            var result = " + (fun.FunctionType.Contains("Task") ? "await" : "") + " _repository." + fun.Name + "(" + Cast(fun.ClassParameters, false) + ");\r\n");
                    insertCode.Insert("\r\n");
                    insertCode.Insert("            return result;\r\n");
                    insertCode.Insert("        }\r\n");
                    insertCode.Insert("\r\n");


                    //返回IResultModel对象
                    //insertCode.Insert("\r\n");
                    //insertCode.Insert("        " + fun.Access + " " + (fun.FunctionType.Contains("Task") ? "async" : "") + " Task<IResultModel> " + fun.Name + "(" + fun.ClassParametersCast + ")\r\n");
                    //insertCode.Insert("        {\r\n");
                    //insertCode.Insert("            var result = " + (fun.FunctionType.Contains("Task") ? "await" : "") + " _repository." + fun.Name + "(" + Cast(fun.ClassParameters, false) + ");\r\n");
                    //insertCode.Insert("            if (result == null)\r\n");
                    //insertCode.Insert("                return ResultModel.NotExists;\r\n");
                    //insertCode.Insert("\r\n");
                    //insertCode.Insert("            return ResultModel.Success(result);\r\n");
                    //insertCode.Insert("        }\r\n");
                    //insertCode.Insert("\r\n");

                }

                projectItem.Save();
            }
        }

        #endregion

        #region 辅助方法


        /// <summary>
        /// 获取class类的基本结构
        /// </summary>
        /// <param name="applicationStr"></param>
        /// <param name="name"></param>
        /// <param name="dirName"></param>
        /// <param name="codeClass"></param>
        /// <returns></returns>
        private ClassModel GetClassModel(string applicationStr, string name, string cnName, string description, string dirName, CodeClass codeClass)
        {
            var model = new ClassModel()
            {
                Namespace = applicationStr,
                Name = name,
                CnName = cnName,
                Description = description,
                DirName = dirName.Replace("\\", ".")
            };

            List<ClassFunction> classFunctions = new List<ClassFunction>();
            List<ClassParameter> classParameters = null;

            var codeMembers = codeClass.Members;
            foreach (CodeElement codeMember in codeMembers)
            {
                if (codeMember.Kind == vsCMElement.vsCMElementFunction)
                {
                    ClassFunction classFunction = new ClassFunction();
                    classParameters = new List<ClassParameter>();

                    CodeFunction function = codeMember as CodeFunction;

                    if (name == function.Name)
                        continue;

                    classFunction.Name = function.Name;
                    classFunction.FunctionType = SimplifyType(function.Type.CodeType.FullName);


                    switch (function.Access)
                    {
                        case vsCMAccess.vsCMAccessPublic:
                            classFunction.Access = "public";
                            break;
                        case vsCMAccess.vsCMAccessPrivate:
                            classFunction.Access = "private";
                            break;
                        default:
                            classFunction.Access = "protected";
                            break;
                    }

                    //获取参数特性
                    foreach (CodeParameter codeParameter in function.Parameters)
                    {
                        ClassParameter classParameter = new ClassParameter();
                        //获取参数名称
                        classParameter.Name = codeParameter.FullName;
                        //获取参数类型
                        classParameter.ParameterType = codeParameter.Type.AsString.Substring(codeParameter.Type.AsString.LastIndexOf(".") + 1);

                        //codeAttribute.Value.Replace("Name = ", "").Replace("\"", "");
                        classParameters.Add(classParameter);
                    }

                    classFunction.ClassParameters = classParameters;
                    classFunction.ClassParametersCast = Cast(classParameters);

                    classFunctions.Add(classFunction);
                }
            }

            model.ClassFunctions = classFunctions;

            return model;
        }

        /// <summary>
        /// 获取class类的基本结构
        /// </summary>
        /// <param name="applicationStr"></param>
        /// <param name="name"></param>
        /// <param name="dirName"></param>
        /// <param name="codeClass"></param>
        /// <returns></returns>
        private ClassModel GetInterfaceModel(string applicationStr, string name, string cnName, string description, string dirName, CodeInterface codeInterface)
        {
            var model = new ClassModel()
            {
                Namespace = applicationStr,
                Name = name,
                CnName = cnName,
                Description = description,
                DirName = dirName.Replace("\\", ".")
            };

            List<ClassFunction> classFunctions = new List<ClassFunction>();
            List<ClassParameter> classParameters = null;

            var codeMembers = codeInterface.Members;
            foreach (CodeElement codeMember in codeMembers)
            {
                if (codeMember.Kind == vsCMElement.vsCMElementFunction)
                {
                    ClassFunction classFunction = new ClassFunction();
                    classParameters = new List<ClassParameter>();

                    CodeFunction function = codeMember as CodeFunction;

                    if (name == function.Name)
                        continue;

                    classFunction.Name = function.Name;
                    classFunction.FunctionType = SimplifyType(function.Type.CodeType.FullName);

                    //获取参数特性
                    foreach (CodeParameter codeParameter in function.Parameters)
                    {
                        ClassParameter classParameter = new ClassParameter();

                        //获取参数名称
                        classParameter.Name = codeParameter.FullName;
                        //获取参数类型
                        classParameter.ParameterType = codeParameter.Type.AsString.Substring(codeParameter.Type.AsString.LastIndexOf(".") + 1);

                        //codeAttribute.Value.Replace("Name = ", "").Replace("\"", "");
                        classParameters.Add(classParameter);
                    }

                    classFunction.ClassParameters = classParameters;
                    classFunction.ClassParametersCast = Cast(classParameters);

                    classFunctions.Add(classFunction);
                }
            }

            model.ClassFunctions = classFunctions;

            return model;
        }



        /// <summary>
        /// 获取解决方案里面所有项目
        /// </summary>
        /// <param name="solution"></param>
        /// <returns></returns>
        private List<Project> GetSolutionProjects(Solution solution)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            List<Project> projectlist = new List<Project>();

            foreach (var temp in solution.Projects)
            {
                if (temp is Project)
                {
                    if (((Project)temp).Kind == ProjectKinds.vsProjectKindSolutionFolder)
                    {
                        projectlist.AddRange(GetSolutionFolderProjects((Project)temp));
                    }
                    else
                    {
                        projectlist.Add((Project)temp);
                    }
                }

            }
            return projectlist;
        }

        private static List<Project> GetSolutionFolderProjects(Project solutionFolder)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            List<Project> list = new List<Project>();
            for (var i = 1; i <= solutionFolder.ProjectItems.Count; i++)
            {
                var subProject = solutionFolder.ProjectItems.Item(i).SubProject;
                if (subProject == null)
                {
                    continue;
                }

                // If this is another solution folder, do a recursive call, otherwise add
                if (subProject.Kind == ProjectKinds.vsProjectKindSolutionFolder)
                {
                    list.AddRange(GetSolutionFolderProjects(subProject));
                }
                else
                {
                    list.Add(subProject);
                }
            }

            return list;
        }


        /// <summary>
        /// 获取类
        /// </summary>
        /// <param name="codeElements"></param>
        /// <returns></returns>
        private CodeClass GetClass(CodeElements codeElements)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var elements = codeElements.Cast<CodeElement>().ToList();
            var result = elements.FirstOrDefault(codeElement => codeElement.Kind == vsCMElement.vsCMElementClass) as CodeClass;
            if (result != null)
            {
                return result;
            }
            foreach (var codeElement in elements)
            {
                result = GetClass(codeElement?.Children);
                if (result != null)
                {
                    return result;
                }
            }
            return null;
        }

        /// <summary>
        /// 获取函数
        /// </summary>
        /// <param name="codeElements"></param>
        /// <returns></returns>
        private CodeFunction GetFunction(CodeElements codeElements)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var elements = codeElements.Cast<CodeElement>().ToList();
            var result = elements.FirstOrDefault(codeElement => codeElement.Kind == vsCMElement.vsCMElementFunction) as CodeFunction;
            if (result != null)
            {
                return result;
            }
            foreach (var codeElement in elements)
            {
                result = GetFunction(codeElement.Children);
                if (result != null)
                {
                    return result;
                }
            }
            return null;
        }

        /// <summary>
        /// 获取接口
        /// </summary>
        /// <param name="codeElements"></param>
        /// <returns></returns>
        private CodeInterface GetInterface(CodeElements codeElements)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var elements = codeElements.Cast<CodeElement>().ToList();
            var result = elements.FirstOrDefault(codeElement => codeElement.Kind == vsCMElement.vsCMElementInterface) as CodeInterface;
            if (result != null)
            {
                return result;
            }
            foreach (var codeElement in elements)
            {
                result = GetInterface(codeElement?.Children);
                if (result != null)
                {
                    return result;
                }
            }
            return null;
        }


        /// <summary>
        /// 获取当前所选文件去除项目目录后的文件夹结构
        /// </summary>
        /// <param name="selectProjectItem"></param>
        /// <returns></returns>
        private string GetSelectFileDirPath(Project topProject, ProjectItem selectProjectItem)
        {
            string dirPath = "";
            if (selectProjectItem != null)
            {
                //所选文件对应的路径
                string fileNames = selectProjectItem.FileNames[0];
                string selectedFullName = fileNames.Substring(0, fileNames.LastIndexOf('\\'));

                //所选文件所在的项目
                if (topProject != null)
                {
                    //项目目录
                    string projectFullName = topProject.FullName.Substring(0, topProject.FullName.LastIndexOf('\\'));

                    //当前所选文件去除项目目录后的文件夹结构
                    dirPath = selectedFullName.Replace(projectFullName, "");
                }
            }

            return dirPath.Substring(1);
        }

        /// <summary>
        /// 添加文件到项目中
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="content"></param>
        /// <param name="fileName"></param>
        private void AddFileToProjectItem(ProjectItem folder, string content, string fileName)
        {
            try
            {
                string path = Path.GetTempPath();
                Directory.CreateDirectory(path);
                string file = Path.Combine(path, fileName);
                File.WriteAllText(file, content, System.Text.Encoding.UTF8);
                try
                {
                    folder.ProjectItems.AddFromFileCopy(file);
                }
                finally
                {
                    File.Delete(file);
                }
            }
            catch (Exception ex)
            {

            }
        }

        /// <summary>
        /// 添加文件到指定目录
        /// </summary>
        /// <param name="directoryPathOrFullPath"></param>
        /// <param name="content"></param>
        /// <param name="fileName"></param>
        private void AddFileToDirectory(string directoryPathOrFullPath, string content, string fileName = "")
        {
            try
            {
                string file = string.IsNullOrEmpty(fileName) ? directoryPathOrFullPath : Path.Combine(directoryPathOrFullPath, fileName);
                File.WriteAllText(file, content, System.Text.Encoding.UTF8);
            }
            catch (Exception ex)
            {

            }
        }

        /// <summary>
        /// 拼接参数
        /// </summary>
        /// <param name="classParameters"></param>
        /// <returns></returns>
        private string Cast(List<ClassParameter> classParameters, bool isDefine = true)
        {
            string res = "";
            if (isDefine)
                classParameters.ForEach(m => res += ", " + m.ParameterType + " " + m.Name);
            else
                classParameters.ForEach(m => res += ", " + m.Name);

            if (res != "")
                return res.Substring(2);
            return res;
        }

        /// <summary>
        /// 优化类型名称
        /// </summary>
        /// <param name="codeType"></param>
        /// <returns></returns>
        private string SimplifyType(string codeType, char c = '<')
        {
            string res = "";
            var codeTypes = codeType.Split(new char[] { c });
            foreach (string s in codeTypes)
            {
                if (s.Contains(","))
                {
                    res += SimplifyType(s, ',') + c;
                    continue;
                }

                if (s.LastIndexOf(".") >= 0)
                    res += s.Substring(s.LastIndexOf(".") + 1) + c;
                else
                    res += s + c;
            }
            if (codeType.Length != res.Length)
                res = res.Substring(0, res.Length - 1);


            return res;
        }


        #endregion


    }
}
