using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using GrasshopperMCP.Models;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Special;
using Grasshopper.Kernel.Parameters;
using Rhino;
using Grasshopper;
using System.Reflection;

namespace GrasshopperMCP.Commands
{
    /// <summary>
    /// 處理代碼執行相關命令的處理器
    /// </summary>
    public static class CodeExecutionCommandHandler
    {
        /// <summary>
        /// 執行代碼命令
        /// </summary>
        /// <param name="command">包含代碼和執行參數的命令</param>
        /// <returns>執行結果</returns>
        public static object ExecuteCode(Command command)
        {
            string code = command.GetParameter<string>("code");
            string language = command.GetParameter<string>("language"); // "python" or "csharp"
            double x = command.GetParameter<double>("x");
            double y = command.GetParameter<double>("y");
            
            // 獲取可選的輸入和輸出參數配置
            var inputParams = command.Parameters.ContainsKey("inputs") ? 
                command.Parameters["inputs"] as List<object> : new List<object>();
            var outputParams = command.Parameters.ContainsKey("outputs") ? 
                command.Parameters["outputs"] as List<object> : new List<object>();
            
            if (string.IsNullOrEmpty(code))
            {
                throw new ArgumentException("Code is required");
            }
            
            // 確保語言識別
            if (string.IsNullOrEmpty(language))
            {
                language = "python"; // 預設為 Python
            }
            
            // 標準化語言名稱
            language = language.ToLowerInvariant();
            if (language == "csharp" || language == "c#")
            {
                language = "csharp";
            }
            else if (language == "python" || language == "python3")
            {
                language = "python";
            }
            
            // 添加語言指定符到代碼開頭（如果沒有的話）
            string processedCode = AddLanguageSpecifier(code, language);
            
            object result = null;
            Exception exception = null;
            
            // 在 UI 線程上執行
            RhinoApp.InvokeOnUiThread(new Action(() =>
            {
                try
                {
                    // 獲取 Grasshopper 文檔
                    var doc = Grasshopper.Instances.ActiveCanvas?.Document;
                    if (doc == null)
                    {
                        throw new InvalidOperationException("No active Grasshopper document");
                    }
                    
                    // 創建腳本組件
                    IGH_DocumentObject scriptComponent = CreateScriptComponent(processedCode, language, inputParams, outputParams);
                    
                    if (scriptComponent != null)
                    {
                        // 確保組件有有效的屬性對象
                        if (scriptComponent.Attributes == null)
                        {
                            scriptComponent.CreateAttributes();
                        }
                        
                        // 設置位置
                        scriptComponent.Attributes.Pivot = new System.Drawing.PointF((float)x, (float)y);
                        
                        // 添加到文檔
                        doc.AddObject(scriptComponent, false);
                        
                        // 刷新畫布
                        doc.NewSolution(false);
                        
                        // 返回組件信息
                        result = new
                        {
                            id = scriptComponent.InstanceGuid.ToString(),
                            type = scriptComponent.GetType().Name,
                            name = scriptComponent.NickName,
                            language = language,
                            x = scriptComponent.Attributes.Pivot.X,
                            y = scriptComponent.Attributes.Pivot.Y,
                            message = "Script component created successfully with language detection",
                            parameters = new
                            {
                                inputs = inputParams?.Count ?? 0,
                                outputs = outputParams?.Count ?? 0
                            }
                        };
                    }
                    else
                    {
                        throw new InvalidOperationException("Failed to create script component");
                    }
                }
                catch (Exception ex)
                {
                    exception = ex;
                    RhinoApp.WriteLine($"Error in ExecuteCode: {ex.Message}");
                }
            }));
            
            // 等待 UI 線程操作完成
            while (result == null && exception == null)
            {
                Thread.Sleep(10);
            }
            
            // 如果有異常，拋出
            if (exception != null)
            {
                throw exception;
            }

            return Response.Ok(new { result });
        }
        
        /// <summary>
        /// 添加語言指定符到代碼開頭
        /// </summary>
        private static string AddLanguageSpecifier(string code, string language)
        {
            if (string.IsNullOrEmpty(code))
                return code;
            
            // 檢查代碼是否已經包含語言指定符
            var lines = code.Split('\n');
            var firstLine = lines.Length > 0 ? lines[0].Trim() : "";
            
            // 如果已經有語言指定符，返回原代碼
            if (firstLine.StartsWith("#!") || firstLine.StartsWith("//"))
            {
                return code;
            }
            
            // 添加語言指定符
            string specifier = "";
            if (language == "python")
            {
                specifier = "#! python 3\n";
            }
            else if (language == "csharp")
            {
                specifier = "// #! csharp\n";
            }
            
            return specifier + code;
        }
        
        /// <summary>
        /// 創建腳本組件
        /// </summary>
        private static IGH_DocumentObject CreateScriptComponent(string code, string language, List<object> inputParams, List<object> outputParams)
        {
            try
            {
                // 首先查找新的統一Script組件
                var scriptProxy = Grasshopper.Instances.ComponentServer.ObjectProxies
                    .FirstOrDefault(p => p.Desc.Name.Equals("Script", StringComparison.OrdinalIgnoreCase));
                
                if (scriptProxy != null)
                {
                    RhinoApp.WriteLine($"Creating unified Script component with language: {language}");
                    var scriptComponent = scriptProxy.CreateInstance();
                    
                    // 設置語言
                    SetScriptLanguage(scriptComponent, language);
                    
                    // 設置代碼
                    SetScriptCode(scriptComponent, code);
                    
                    // 配置參數
                    ConfigureScriptParameters(scriptComponent, inputParams, outputParams);
                    
                    return scriptComponent;
                }
                else
                {
                    // 回退到特定語言的組件
                    return CreateLanguageSpecificComponent(code, language, inputParams, outputParams);
                }
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Error creating script component: {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// 創建特定語言的腳本組件
        /// </summary>
        private static IGH_DocumentObject CreateLanguageSpecificComponent(string code, string language, List<object> inputParams, List<object> outputParams)
        {
            if (language == "python")
            {
                return CreatePython3Component(code, inputParams, outputParams);
            }
            else if (language == "csharp")
            {
                return CreateCSharpComponent(code, inputParams, outputParams);
            }
            else
            {
                throw new ArgumentException($"Unsupported language: {language}");
            }
        }
        
        /// <summary>
        /// 配置腳本參數
        /// </summary>
        private static void ConfigureScriptParameters(IGH_DocumentObject component, List<object> inputParams, List<object> outputParams)
        {
            try
            {
                // 記錄參數指導信息
                if (inputParams != null && inputParams.Count > 0)
                {
                    RhinoApp.WriteLine("=== INPUT PARAMETERS GUIDANCE ===");
                    RhinoApp.WriteLine("After creating the script component, manually configure these input parameters:");
                    
                    foreach (var paramConfig in inputParams)
                    {
                        var paramDict = paramConfig as Dictionary<string, object>;
                        if (paramDict != null)
                        {
                            var paramName = paramDict.ContainsKey("name") ? paramDict["name"].ToString() : "x";
                            var paramType = paramDict.ContainsKey("type") ? paramDict["type"].ToString() : "number";
                            var paramDescription = paramDict.ContainsKey("description") ? paramDict["description"].ToString() : "";
                            
                            RhinoApp.WriteLine($"  • Add input parameter: '{paramName}' (type: {paramType})");
                            if (!string.IsNullOrEmpty(paramDescription))
                            {
                                RhinoApp.WriteLine($"    Description: {paramDescription}");
                            }
                        }
                    }
                }
                
                if (outputParams != null && outputParams.Count > 0)
                {
                    RhinoApp.WriteLine("=== OUTPUT PARAMETERS GUIDANCE ===");
                    RhinoApp.WriteLine("After creating the script component, manually configure these output parameters:");
                    
                    foreach (var paramConfig in outputParams)
                    {
                        var paramDict = paramConfig as Dictionary<string, object>;
                        if (paramDict != null)
                        {
                            var paramName = paramDict.ContainsKey("name") ? paramDict["name"].ToString() : "a";
                            var paramType = paramDict.ContainsKey("type") ? paramDict["type"].ToString() : "geometry";
                            var paramDescription = paramDict.ContainsKey("description") ? paramDict["description"].ToString() : "";
                            
                            RhinoApp.WriteLine($"  • Add output parameter: '{paramName}' (type: {paramType})");
                            if (!string.IsNullOrEmpty(paramDescription))
                            {
                                RhinoApp.WriteLine($"    Description: {paramDescription}");
                            }
                        }
                    }
                }
                
                RhinoApp.WriteLine("=== MANUAL SETUP INSTRUCTIONS ===");
                RhinoApp.WriteLine("1. Zoom in on the script component until you see ⊕ and ⊖ buttons");
                RhinoApp.WriteLine("2. Use ⊕ to add input/output parameters");
                RhinoApp.WriteLine("3. Right-click on parameters to set name and type hints");
                RhinoApp.WriteLine("4. Use the variable names in your script code");
                RhinoApp.WriteLine("=====================================");
                
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Error configuring script parameters: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 創建 Python 3 腳本組件 (使用 Rhino 8 SR18 新 API)
        /// </summary>
        /// <param name="code">Python 代碼</param>
        /// <param name="inputParams">輸入參數配置</param>
        /// <param name="outputParams">輸出參數配置</param>
        /// <returns>Python 3 腳本組件</returns>
        private static IGH_DocumentObject CreatePython3Component(string code, List<object> inputParams, List<object> outputParams)
        {
            try
            {
                // 首先嘗試使用 Rhino 8 SR18 新 API
                var component = TryCreateWithNewAPI("Python3Component", code, inputParams, outputParams);
                if (component != null)
                {
                    return component;
                }
                
                // 回退到舊的 API - 優先尋找新的Script組件
                var scriptProxy = Grasshopper.Instances.ComponentServer.ObjectProxies
                    .FirstOrDefault(p => p.Desc.Name.Equals("Script", StringComparison.OrdinalIgnoreCase));
                
                if (scriptProxy != null)
                {
                    RhinoApp.WriteLine("Found new Script component, creating instance");
                    var scriptComponent = scriptProxy.CreateInstance();
                    
                    // 設置語言為Python 3
                    SetScriptLanguage(scriptComponent, "Python 3");
                    
                    // 設置代碼
                    SetScriptCode(scriptComponent, code);
                    
                    // 配置輸入和輸出參數
                    ConfigureParametersLegacy(scriptComponent, inputParams, outputParams);
                    
                    return scriptComponent;
                }
                else
                {
                    // 如果沒有找到新的Script組件，嘗試舊的Python組件
                    var pythonProxy = Grasshopper.Instances.ComponentServer.ObjectProxies
                        .FirstOrDefault(p => p.Desc.Name.Contains("Python") && p.Desc.Name.Contains("Script"));
                    
                    if (pythonProxy != null)
                    {
                        RhinoApp.WriteLine("Using legacy Python component");
                        var legacyComponent = pythonProxy.CreateInstance();
                        
                        // 設置代碼
                        SetScriptCode(legacyComponent, code);
                        
                        // 配置輸入和輸出參數
                        ConfigureParametersLegacy(legacyComponent, inputParams, outputParams);
                        
                        return legacyComponent;
                    }
                    else
                    {
                        throw new InvalidOperationException("Script component not found. Make sure scripting is available.");
                    }
                }
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Error creating Python script component: {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// 創建 C# 腳本組件 (使用 Rhino 8 SR18 新 API)
        /// </summary>
        /// <param name="code">C# 代碼</param>
        /// <param name="inputParams">輸入參數配置</param>
        /// <param name="outputParams">輸出參數配置</param>
        /// <returns>C# 腳本組件</returns>
        private static IGH_DocumentObject CreateCSharpComponent(string code, List<object> inputParams, List<object> outputParams)
        {
            try
            {
                // 首先嘗試使用 Rhino 8 SR18 新 API
                var component = TryCreateWithNewAPI("CSharpComponent", code, inputParams, outputParams);
                if (component != null)
                {
                    return component;
                }
                
                // 回退到舊的 API - 優先尋找新的Script組件
                var scriptProxy = Grasshopper.Instances.ComponentServer.ObjectProxies
                    .FirstOrDefault(p => p.Desc.Name.Equals("Script", StringComparison.OrdinalIgnoreCase));
                
                if (scriptProxy != null)
                {
                    RhinoApp.WriteLine("Found new Script component, creating instance");
                    var scriptComponent = scriptProxy.CreateInstance();
                    
                    // 設置語言為C#
                    SetScriptLanguage(scriptComponent, "C#");
                    
                    // 設置代碼
                    SetScriptCode(scriptComponent, code);
                    
                    // 配置輸入和輸出參數
                    ConfigureParametersLegacy(scriptComponent, inputParams, outputParams);
                    
                    return scriptComponent;
                }
                else
                {
                    // 如果沒有找到新的Script組件，嘗試舊的C#組件
                    var csharpProxy = Grasshopper.Instances.ComponentServer.ObjectProxies
                        .FirstOrDefault(p => p.Desc.Name.Contains("C#") && p.Desc.Name.Contains("Script"));
                    
                    if (csharpProxy != null)
                    {
                        RhinoApp.WriteLine("Using legacy C# component");
                        var legacyComponent = csharpProxy.CreateInstance();
                        
                        // 設置代碼
                        SetScriptCode(legacyComponent, code);
                        
                        // 配置輸入和輸出參數
                        ConfigureParametersLegacy(legacyComponent, inputParams, outputParams);
                        
                        return legacyComponent;
                    }
                    else
                    {
                        throw new InvalidOperationException("Script component not found. Make sure scripting is available.");
                    }
                }
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Error creating C# script component: {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// 嘗試使用 Rhino 8 SR18 新 API 創建組件
        /// </summary>
        /// <param name="componentTypeName">組件類型名稱</param>
        /// <param name="code">代碼</param>
        /// <param name="inputParams">輸入參數配置</param>
        /// <param name="outputParams">輸出參數配置</param>
        /// <returns>創建的組件或 null</returns>
        private static IGH_DocumentObject TryCreateWithNewAPI(string componentTypeName, string code, List<object> inputParams, List<object> outputParams)
        {
            try
            {
                // 加載 RhinoCodePluginGH.rhp
                try
                {
                    var pluginPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), 
                        "Rhino 8", "Plug-ins", "RhinoCodePluginGH.rhp");
                    
                    if (System.IO.File.Exists(pluginPath))
                    {
                        var pluginAssembly = Assembly.LoadFrom(pluginPath);
                        RhinoApp.WriteLine($"Successfully loaded RhinoCodePluginGH.rhp from {pluginPath}");
                    }
                }
                catch (Exception loadEx)
                {
                    RhinoApp.WriteLine($"Failed to load RhinoCodePluginGH.rhp: {loadEx.Message}");
                }
                
                // 嘗試使用 Rhino 8 SR18 的新 API
                var assemblies = AppDomain.CurrentDomain.GetAssemblies();
                var rhinoCodeAssembly = assemblies.FirstOrDefault(a => a.GetName().Name == "RhinoCodePluginGH");
                
                RhinoApp.WriteLine($"Looking for RhinoCodePluginGH assembly... Found: {rhinoCodeAssembly != null}");
                
                if (rhinoCodeAssembly != null)
                {
                    var componentType = rhinoCodeAssembly.GetType($"RhinoCodePluginGH.Components.{componentTypeName}");
                    RhinoApp.WriteLine($"Looking for type RhinoCodePluginGH.Components.{componentTypeName}... Found: {componentType != null}");
                    
                    if (componentType != null)
                    {
                        // 使用新的 Create 方法
                        var createMethod = componentType.GetMethod("Create", BindingFlags.Public | BindingFlags.Static);
                        RhinoApp.WriteLine($"Looking for static Create method... Found: {createMethod != null}");
                        
                        if (createMethod != null)
                        {
                            object component = null;
                            var parameters = createMethod.GetParameters();
                            RhinoApp.WriteLine($"Create method has {parameters.Length} parameters");
                            
                            if (parameters.Length == 2)
                            {
                                // Create(string name, string code)
                                component = createMethod.Invoke(null, new object[] { "MCP Script", code });
                                RhinoApp.WriteLine("Used Create(name, code) method");
                            }
                            else if (parameters.Length == 1)
                            {
                                // Create(string name)
                                component = createMethod.Invoke(null, new object[] { "MCP Script" });
                                RhinoApp.WriteLine("Used Create(name) method");
                                
                                // 設置源碼
                                var setSourceMethod = componentType.GetMethod("SetSource", BindingFlags.Public | BindingFlags.Instance);
                                if (setSourceMethod != null)
                                {
                                    setSourceMethod.Invoke(component, new object[] { code });
                                    RhinoApp.WriteLine("Set source code using SetSource method");
                                }
                            }
                            else if (parameters.Length == 0)
                            {
                                // Create()
                                component = createMethod.Invoke(null, new object[] { });
                                RhinoApp.WriteLine("Used Create() method");
                                
                                // 設置名稱
                                var nameProperty = componentType.GetProperty("NickName", BindingFlags.Public | BindingFlags.Instance);
                                if (nameProperty != null && nameProperty.CanWrite)
                                {
                                    nameProperty.SetValue(component, "MCP Script");
                                    RhinoApp.WriteLine("Set component name to 'MCP Script'");
                                }
                                
                                // 設置源碼
                                var setSourceMethod = componentType.GetMethod("SetSource", BindingFlags.Public | BindingFlags.Instance);
                                if (setSourceMethod != null)
                                {
                                    setSourceMethod.Invoke(component, new object[] { code });
                                    RhinoApp.WriteLine("Set source code using SetSource method");
                                }
                            }
                            
                            if (component != null)
                            {
                                RhinoApp.WriteLine($"Successfully created component of type {componentType.Name}");
                                
                                // 配置參數
                                ConfigureParametersNewAPI(component, inputParams, outputParams);
                                
                                return component as IGH_DocumentObject;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"New API not available, falling back to legacy API: {ex.Message}");
                RhinoApp.WriteLine($"Stack trace: {ex.StackTrace}");
            }
            
            return null;
        }
        
        /// <summary>
        /// 使用新 API 配置參數
        /// </summary>
        /// <param name="component">組件</param>
        /// <param name="inputParams">輸入參數配置</param>
        /// <param name="outputParams">輸出參數配置</param>
        private static void ConfigureParametersNewAPI(object component, List<object> inputParams, List<object> outputParams)
        {
            try
            {
                var componentType = component.GetType();
                RhinoApp.WriteLine($"Configuring parameters for component type: {componentType.Name}");
                
                // 獲取 ScriptVariableParam 類型
                var assemblies = AppDomain.CurrentDomain.GetAssemblies();
                var rhinoCodeAssembly = assemblies.FirstOrDefault(a => a.GetName().Name == "RhinoCodePluginGH");
                var scriptVariableParamType = rhinoCodeAssembly?.GetType("RhinoCodePluginGH.Parameters.ScriptVariableParam");
                
                RhinoApp.WriteLine($"ScriptVariableParam type found: {scriptVariableParamType != null}");
                
                if (scriptVariableParamType != null)
                {
                    // 配置輸入參數
                    ConfigureInputParametersNewAPI(component, inputParams, scriptVariableParamType);
                    
                    // 配置輸出參數
                    ConfigureOutputParametersNewAPI(component, outputParams, scriptVariableParamType);
                    
                    // 調用 VariableParameterMaintenance 方法
                    var maintenanceMethod = componentType.GetMethod("VariableParameterMaintenance", BindingFlags.Public | BindingFlags.Instance);
                    if (maintenanceMethod != null)
                    {
                        maintenanceMethod.Invoke(component, null);
                        RhinoApp.WriteLine("Called VariableParameterMaintenance method");
                    }
                    else
                    {
                        RhinoApp.WriteLine("VariableParameterMaintenance method not found");
                    }
                    
                    // 調用 SetParametersToScript 方法來同步參數到腳本簽名
                    var setParametersToScriptMethod = componentType.GetMethod("SetParametersToScript", BindingFlags.Public | BindingFlags.Instance);
                    if (setParametersToScriptMethod != null)
                    {
                        setParametersToScriptMethod.Invoke(component, null);
                        RhinoApp.WriteLine("Called SetParametersToScript method - parameters synchronized to script signature");
                    }
                    else
                    {
                        RhinoApp.WriteLine("SetParametersToScript method not found");
                    }
                }
                else
                {
                    RhinoApp.WriteLine("ScriptVariableParam type not found, falling back to legacy parameter configuration");
                }
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Error configuring parameters with new API: {ex.Message}");
                RhinoApp.WriteLine($"Stack trace: {ex.StackTrace}");
            }
        }
        
        /// <summary>
        /// 使用新 API 配置輸入參數
        /// </summary>
        /// <param name="component">組件</param>
        /// <param name="inputParams">輸入參數配置</param>
        /// <param name="scriptVariableParamType">ScriptVariableParam 類型</param>
        private static void ConfigureInputParametersNewAPI(object component, List<object> inputParams, Type scriptVariableParamType)
        {
            if (inputParams == null || inputParams.Count == 0)
            {
                RhinoApp.WriteLine("No input parameters to configure");
                return;
            }
            
            try
            {
                var componentType = component.GetType();
                var paramsProperty = componentType.GetProperty("Params");
                
                if (paramsProperty != null)
                {
                    var paramsObject = paramsProperty.GetValue(component);
                    var paramsType = paramsObject.GetType();
                    var registerInputMethod = paramsType.GetMethod("RegisterInputParam", BindingFlags.Public | BindingFlags.Instance);
                    
                    RhinoApp.WriteLine($"RegisterInputParam method found: {registerInputMethod != null}");
                    
                    if (registerInputMethod != null)
                    {
                        foreach (var paramConfig in inputParams)
                        {
                            var paramDict = paramConfig as Dictionary<string, object>;
                            if (paramDict != null)
                            {
                                var paramName = paramDict.ContainsKey("name") ? paramDict["name"].ToString() : "x";
                                var paramType = paramDict.ContainsKey("type") ? paramDict["type"].ToString() : "number";
                                var paramDescription = paramDict.ContainsKey("description") ? paramDict["description"].ToString() : "Input parameter";
                                var paramOptional = paramDict.ContainsKey("optional") ? Convert.ToBoolean(paramDict["optional"]) : true;
                                
                                RhinoApp.WriteLine($"Configuring input parameter: {paramName} ({paramType})");
                                
                                // 創建 ScriptVariableParam 實例
                                var param = Activator.CreateInstance(scriptVariableParamType, paramName);
                                
                                // 設置屬性
                                var prettyNameProperty = scriptVariableParamType.GetProperty("PrettyName");
                                if (prettyNameProperty != null)
                                {
                                    prettyNameProperty.SetValue(param, paramName);
                                }
                                
                                var toolTipProperty = scriptVariableParamType.GetProperty("ToolTip");
                                if (toolTipProperty != null)
                                {
                                    toolTipProperty.SetValue(param, paramDescription);
                                }
                                
                                var optionalProperty = scriptVariableParamType.GetProperty("Optional");
                                if (optionalProperty != null)
                                {
                                    optionalProperty.SetValue(param, paramOptional);
                                }
                                
                                // 設置類型提示
                                var typeHintsProperty = scriptVariableParamType.GetProperty("TypeHints");
                                if (typeHintsProperty != null)
                                {
                                    var typeHints = typeHintsProperty.GetValue(param);
                                    var selectMethod = typeHints.GetType().GetMethod("Select", new[] { typeof(string) });
                                    if (selectMethod != null)
                                    {
                                        var typeHintString = GetTypeHintString(paramType);
                                        selectMethod.Invoke(typeHints, new object[] { typeHintString });
                                        RhinoApp.WriteLine($"Set type hint to: {typeHintString}");
                                    }
                                }
                                
                                // 創建屬性
                                var createAttributesMethod = scriptVariableParamType.GetMethod("CreateAttributes");
                                if (createAttributesMethod != null)
                                {
                                    createAttributesMethod.Invoke(param, null);
                                }
                                
                                // 註冊參數
                                registerInputMethod.Invoke(paramsObject, new object[] { param });
                                RhinoApp.WriteLine($"Successfully registered input parameter: {paramName}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Error configuring input parameters with new API: {ex.Message}");
                RhinoApp.WriteLine($"Stack trace: {ex.StackTrace}");
            }
        }
        
        /// <summary>
        /// 使用新 API 配置輸出參數
        /// </summary>
        /// <param name="component">組件</param>
        /// <param name="outputParams">輸出參數配置</param>
        /// <param name="scriptVariableParamType">ScriptVariableParam 類型</param>
        private static void ConfigureOutputParametersNewAPI(object component, List<object> outputParams, Type scriptVariableParamType)
        {
            if (outputParams == null || outputParams.Count == 0)
            {
                RhinoApp.WriteLine("No output parameters to configure");
                return;
            }
            
            try
            {
                var componentType = component.GetType();
                var paramsProperty = componentType.GetProperty("Params");
                
                if (paramsProperty != null)
                {
                    var paramsObject = paramsProperty.GetValue(component);
                    var paramsType = paramsObject.GetType();
                    var registerOutputMethod = paramsType.GetMethod("RegisterOutputParam", BindingFlags.Public | BindingFlags.Instance);
                    
                    RhinoApp.WriteLine($"RegisterOutputParam method found: {registerOutputMethod != null}");
                    
                    if (registerOutputMethod != null)
                    {
                        foreach (var paramConfig in outputParams)
                        {
                            var paramDict = paramConfig as Dictionary<string, object>;
                            if (paramDict != null)
                            {
                                var paramName = paramDict.ContainsKey("name") ? paramDict["name"].ToString() : "a";
                                var paramType = paramDict.ContainsKey("type") ? paramDict["type"].ToString() : "geometry";
                                var paramDescription = paramDict.ContainsKey("description") ? paramDict["description"].ToString() : "Output parameter";
                                
                                RhinoApp.WriteLine($"Configuring output parameter: {paramName} ({paramType})");
                                
                                // 創建 ScriptVariableParam 實例
                                var param = Activator.CreateInstance(scriptVariableParamType, paramName);
                                
                                // 設置屬性
                                var prettyNameProperty = scriptVariableParamType.GetProperty("PrettyName");
                                if (prettyNameProperty != null)
                                {
                                    prettyNameProperty.SetValue(param, paramName);
                                }
                                
                                var toolTipProperty = scriptVariableParamType.GetProperty("ToolTip");
                                if (toolTipProperty != null)
                                {
                                    toolTipProperty.SetValue(param, paramDescription);
                                }
                                
                                // 設置類型提示
                                var typeHintsProperty = scriptVariableParamType.GetProperty("TypeHints");
                                if (typeHintsProperty != null)
                                {
                                    var typeHints = typeHintsProperty.GetValue(param);
                                    var selectMethod = typeHints.GetType().GetMethod("Select", new[] { typeof(string) });
                                    if (selectMethod != null)
                                    {
                                        var typeHintString = GetTypeHintString(paramType);
                                        selectMethod.Invoke(typeHints, new object[] { typeHintString });
                                        RhinoApp.WriteLine($"Set type hint to: {typeHintString}");
                                    }
                                }
                                
                                // 創建屬性
                                var createAttributesMethod = scriptVariableParamType.GetMethod("CreateAttributes");
                                if (createAttributesMethod != null)
                                {
                                    createAttributesMethod.Invoke(param, null);
                                }
                                
                                // 註冊參數
                                registerOutputMethod.Invoke(paramsObject, new object[] { param });
                                RhinoApp.WriteLine($"Successfully registered output parameter: {paramName}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Error configuring output parameters with new API: {ex.Message}");
                RhinoApp.WriteLine($"Stack trace: {ex.StackTrace}");
            }
        }
        
        /// <summary>
        /// 使用傳統 API 配置參數
        /// </summary>
        /// <param name="component">組件</param>
        /// <param name="inputParams">輸入參數配置</param>
        /// <param name="outputParams">輸出參數配置</param>
        private static void ConfigureParametersLegacy(IGH_DocumentObject component, List<object> inputParams, List<object> outputParams)
        {
            try
            {
                var componentType = component.GetType();
                var paramsProperty = componentType.GetProperty("Params");
                
                if (paramsProperty != null)
                {
                    var paramsObject = paramsProperty.GetValue(component);
                    
                    // 配置輸入參數
                    ConfigureInputParametersLegacy(paramsObject, inputParams);
                    
                    // 配置輸出參數
                    ConfigureOutputParametersLegacy(paramsObject, outputParams);
                    
                    // 調用參數變化通知
                    var onParametersChangedMethod = paramsObject.GetType().GetMethod("OnParametersChanged", BindingFlags.Public | BindingFlags.Instance);
                    if (onParametersChangedMethod != null)
                    {
                        onParametersChangedMethod.Invoke(paramsObject, null);
                    }
                }
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Error configuring parameters with legacy API: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 使用傳統 API 配置輸入參數
        /// </summary>
        /// <param name="paramsObject">參數管理對象</param>
        /// <param name="inputParams">輸入參數配置</param>
        private static void ConfigureInputParametersLegacy(object paramsObject, List<object> inputParams)
        {
            if (inputParams == null || inputParams.Count == 0) return;
            
            try
            {
                var paramsType = paramsObject.GetType();
                var registerInputMethod = paramsType.GetMethod("RegisterInputParam", BindingFlags.Public | BindingFlags.Instance);
                
                if (registerInputMethod != null)
                {
                    foreach (var paramConfig in inputParams)
                    {
                        var paramDict = paramConfig as Dictionary<string, object>;
                        if (paramDict != null)
                        {
                            var paramName = paramDict.ContainsKey("name") ? paramDict["name"].ToString() : "x";
                            var paramType = paramDict.ContainsKey("type") ? paramDict["type"].ToString() : "number";
                            var paramDescription = paramDict.ContainsKey("description") ? paramDict["description"].ToString() : "Input parameter";
                            var paramOptional = paramDict.ContainsKey("optional") ? Convert.ToBoolean(paramDict["optional"]) : true;
                            
                            // 創建適當類型的參數
                            var param = CreateParameterByType(paramType, paramName, paramDescription, paramOptional);
                            if (param != null)
                            {
                                registerInputMethod.Invoke(paramsObject, new object[] { param });
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Error configuring input parameters with legacy API: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 使用傳統 API 配置輸出參數
        /// </summary>
        /// <param name="paramsObject">參數管理對象</param>
        /// <param name="outputParams">輸出參數配置</param>
        private static void ConfigureOutputParametersLegacy(object paramsObject, List<object> outputParams)
        {
            if (outputParams == null || outputParams.Count == 0) return;
            
            try
            {
                var paramsType = paramsObject.GetType();
                var registerOutputMethod = paramsType.GetMethod("RegisterOutputParam", BindingFlags.Public | BindingFlags.Instance);
                
                if (registerOutputMethod != null)
                {
                    foreach (var paramConfig in outputParams)
                    {
                        var paramDict = paramConfig as Dictionary<string, object>;
                        if (paramDict != null)
                        {
                            var paramName = paramDict.ContainsKey("name") ? paramDict["name"].ToString() : "a";
                            var paramType = paramDict.ContainsKey("type") ? paramDict["type"].ToString() : "geometry";
                            var paramDescription = paramDict.ContainsKey("description") ? paramDict["description"].ToString() : "Output parameter";
                            
                            // 創建適當類型的參數
                            var param = CreateParameterByType(paramType, paramName, paramDescription, false);
                            if (param != null)
                            {
                                registerOutputMethod.Invoke(paramsObject, new object[] { param });
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Error configuring output parameters with legacy API: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 根據類型創建參數
        /// </summary>
        /// <param name="paramType">參數類型</param>
        /// <param name="name">參數名稱</param>
        /// <param name="description">參數描述</param>
        /// <param name="optional">是否可選</param>
        /// <returns>創建的參數</returns>
        private static IGH_Param CreateParameterByType(string paramType, string name, string description, bool optional)
        {
            IGH_Param param = null;
            
            try
            {
                switch (paramType.ToLower())
                {
                    case "number":
                    case "double":
                    case "float":
                        param = new Param_Number();
                        break;
                        
                    case "integer":
                    case "int":
                        param = new Param_Integer();
                        break;
                        
                    case "string":
                    case "text":
                        param = new Param_String();
                        break;
                        
                    case "boolean":
                    case "bool":
                        param = new Param_Boolean();
                        break;
                        
                    case "point":
                    case "point3d":
                        param = new Param_Point();
                        break;
                        
                    case "vector":
                    case "vector3d":
                        param = new Param_Vector();
                        break;
                        
                    case "plane":
                        param = new Param_Plane();
                        break;
                        
                    case "line":
                        param = new Param_Line();
                        break;
                        
                    case "curve":
                        param = new Param_Curve();
                        break;
                        
                    case "surface":
                        param = new Param_Surface();
                        break;
                        
                    case "brep":
                        param = new Param_Brep();
                        break;
                        
                    case "mesh":
                        param = new Param_Mesh();
                        break;
                        
                    case "geometry":
                        param = new Param_Geometry();
                        break;
                        
                    case "colour":
                    case "color":
                        param = new Param_Colour();
                        break;
                        
                    case "matrix":
                        param = new Param_Matrix();
                        break;
                        
                    default:
                        // 使用通用參數類型
                        param = new Param_GenericObject();
                        break;
                }
                
                if (param != null)
                {
                    param.Name = name;
                    param.NickName = name;
                    param.Description = description;
                    param.Optional = optional;
                    param.CreateAttributes();
                }
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Error creating parameter of type {paramType}: {ex.Message}");
            }
            
            return param;
        }
        
        /// <summary>
        /// 獲取類型提示字符串
        /// </summary>
        /// <param name="paramType">參數類型</param>
        /// <returns>類型提示字符串</returns>
        private static string GetTypeHintString(string paramType)
        {
            switch (paramType.ToLower())
            {
                case "number":
                case "double":
                case "float":
                    return "float";
                    
                case "integer":
                case "int":
                    return "int";
                    
                case "string":
                case "text":
                    return "str";
                    
                case "boolean":
                case "bool":
                    return "bool";
                    
                case "point":
                case "point3d":
                    return "Point3d";
                    
                case "vector":
                case "vector3d":
                    return "Vector3d";
                    
                case "plane":
                    return "Plane";
                    
                case "line":
                    return "Line";
                    
                case "curve":
                    return "Curve";
                    
                case "surface":
                    return "Surface";
                    
                case "brep":
                    return "Brep";
                    
                case "mesh":
                    return "Mesh";
                    
                case "geometry":
                    return "GeometryBase";
                    
                case "colour":
                case "color":
                    return "Color";
                    
                case "matrix":
                    return "Matrix";
                    
                default:
                    return "object";
            }
        }
        
        /// <summary>
        /// 設置腳本語言
        /// </summary>
        /// <param name="component">組件</param>
        /// <param name="language">語言名稱</param>
        private static void SetScriptLanguage(IGH_DocumentObject component, string language)
        {
            try
            {
                var componentType = component.GetType();
                RhinoApp.WriteLine($"Setting script language to: {language}");
                
                // 嘗試使用 SetLanguage 方法
                var setLanguageMethod = componentType.GetMethod("SetLanguage", BindingFlags.Public | BindingFlags.Instance);
                if (setLanguageMethod != null)
                {
                    setLanguageMethod.Invoke(component, new object[] { language });
                    RhinoApp.WriteLine($"Successfully set language using SetLanguage method");
                    return;
                }
                
                // 嘗試使用 Language 屬性
                var languageProperty = componentType.GetProperty("Language");
                if (languageProperty != null && languageProperty.CanWrite)
                {
                    languageProperty.SetValue(component, language);
                    RhinoApp.WriteLine($"Successfully set language using Language property");
                    return;
                }
                
                // 嘗試使用反射設置內部語言狀態
                var fieldsAndProperties = componentType.GetFields(BindingFlags.NonPublic | BindingFlags.Instance)
                    .Cast<MemberInfo>()
                    .Concat(componentType.GetProperties(BindingFlags.NonPublic | BindingFlags.Instance));
                
                foreach (var member in fieldsAndProperties)
                {
                    if (member.Name.Contains("Language") || member.Name.Contains("language"))
                    {
                        try
                        {
                            if (member is FieldInfo field && field.FieldType == typeof(string))
                            {
                                field.SetValue(component, language);
                                RhinoApp.WriteLine($"Successfully set language using field: {field.Name}");
                                return;
                            }
                            else if (member is PropertyInfo prop && prop.PropertyType == typeof(string) && prop.CanWrite)
                            {
                                prop.SetValue(component, language);
                                RhinoApp.WriteLine($"Successfully set language using property: {prop.Name}");
                                return;
                            }
                        }
                        catch (Exception ex)
                        {
                            RhinoApp.WriteLine($"Failed to set language using {member.Name}: {ex.Message}");
                        }
                    }
                }
                
                RhinoApp.WriteLine($"Unable to set script language for component type: {componentType.Name}");
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Error setting script language: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 設置腳本代碼
        /// </summary>
        /// <param name="component">組件</param>
        /// <param name="code">代碼</param>
        private static void SetScriptCode(IGH_DocumentObject component, string code)
        {
            try
            {
                var componentType = component.GetType();
                
                // 嘗試使用 SetSource 方法 (新 API)
                var setSourceMethod = componentType.GetMethod("SetSource", BindingFlags.Public | BindingFlags.Instance);
                if (setSourceMethod != null)
                {
                    setSourceMethod.Invoke(component, new object[] { code });
                    return;
                }
                
                // 嘗試使用 Code 屬性 (舊 API)
                var codeProperty = componentType.GetProperty("Code");
                if (codeProperty != null && codeProperty.CanWrite)
                {
                    codeProperty.SetValue(component, code);
                    return;
                }
                
                // 嘗試使用 ScriptSource 屬性 (C# 組件)
                var scriptSourceProperty = componentType.GetProperty("ScriptSource");
                if (scriptSourceProperty != null)
                {
                    var scriptSource = scriptSourceProperty.GetValue(component);
                    if (scriptSource != null)
                    {
                        var scriptCodeProperty = scriptSource.GetType().GetProperty("ScriptCode");
                        if (scriptCodeProperty != null && scriptCodeProperty.CanWrite)
                        {
                            scriptCodeProperty.SetValue(scriptSource, code);
                            return;
                        }
                    }
                }
                
                RhinoApp.WriteLine($"Unable to set script code for component type: {componentType.Name}");
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Error setting script code: {ex.Message}");
            }
        }
    }
} 