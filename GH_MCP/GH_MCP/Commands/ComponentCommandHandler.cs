using System;
using System.Collections.Generic;
using GrasshopperMCP.Models;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Special;
using Rhino;
using Rhino.Geometry;
using Grasshopper;
using System.Linq;
using Grasshopper.Kernel.Components;
using System.Threading;
using GH_MCP.Utils;

namespace GrasshopperMCP.Commands
{
    /// <summary>
    /// 處理組件相關命令的處理器
    /// </summary>
    public static class ComponentCommandHandler
    {
        /// <summary>
        /// 添加組件
        /// </summary>
        /// <param name="command">包含組件類型和位置的命令</param>
        /// <returns>添加的組件信息</returns>
        public static object AddComponent(Command command)
        {
            string type = command.GetParameter<string>("type");
            double x = command.GetParameter<double>("x");
            double y = command.GetParameter<double>("y");
            
            if (string.IsNullOrEmpty(type))
            {
                throw new ArgumentException("Component type is required");
            }
            
            // 使用模糊匹配獲取標準化的元件名稱
            string normalizedType = FuzzyMatcher.GetClosestComponentName(type);
            
            // 記錄請求信息
            RhinoApp.WriteLine($"AddComponent request: type={type}, normalized={normalizedType}, x={x}, y={y}");
            
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
                    
                    // 創建組件
                    IGH_DocumentObject component = null;
                    
                    // 記錄可用的組件類型（僅在第一次調用時記錄）
                    bool loggedComponentTypes = false;
                    if (!loggedComponentTypes)
                    {
                        var availableTypes = Grasshopper.Instances.ComponentServer.ObjectProxies
                            .Select(p => p.Desc.Name)
                            .OrderBy(n => n)
                            .ToList();
                        
                        RhinoApp.WriteLine($"Available component types: {string.Join(", ", availableTypes.Take(50))}...");
                        loggedComponentTypes = true;
                    }
                    
                    // 根據類型創建不同的組件
                    switch (normalizedType.ToLowerInvariant())
                    {
                        // 平面元件
                        case "xy plane":
                            component = CreateComponentByName("XY Plane");
                            break;
                        case "xz plane":
                            component = CreateComponentByName("XZ Plane");
                            break;
                        case "yz plane":
                            component = CreateComponentByName("YZ Plane");
                            break;
                        case "plane 3pt":
                            component = CreateComponentByName("Plane 3Pt");
                            break;
                            
                        // 基本幾何元件
                        case "box":
                            component = CreateComponentByName("Box");
                            break;
                        case "sphere":
                            component = CreateComponentByName("Sphere");
                            break;
                        case "cylinder":
                            component = CreateComponentByName("Cylinder");
                            break;
                        case "cone":
                            component = CreateComponentByName("Cone");
                            break;
                        case "circle":
                            component = CreateComponentByName("Circle");
                            break;
                        case "rectangle":
                            component = CreateComponentByName("Rectangle");
                            break;
                        case "line":
                            component = CreateComponentByName("Line");
                            break;
                            
                        // 參數元件
                        case "point":
                        case "pt":
                        case "pointparam":
                        case "param_point":
                            component = new Param_Point();
                            break;
                        case "curve":
                        case "crv":
                        case "curveparam":
                        case "param_curve":
                            component = new Param_Curve();
                            break;
                        case "circleparam":
                        case "param_circle":
                            component = new Param_Circle();
                            break;
                        case "lineparam":
                        case "param_line":
                            component = new Param_Line();
                            break;
                        case "panel":
                        case "gh_panel":
                            component = new GH_Panel();
                            break;
                        case "slider":
                        case "numberslider":
                        case "gh_numberslider":
                            var slider = new GH_NumberSlider();
                            slider.CreateAttributes(); // 创建属性，确保slider不会崩溃
                            
                            // 檢查是否有自定義範圍參數
                            double minValue = command.Parameters.ContainsKey("min") ? 
                                Convert.ToDouble(command.Parameters["min"]) : 0.0;
                            double maxValue = command.Parameters.ContainsKey("max") ? 
                                Convert.ToDouble(command.Parameters["max"]) : 10.0;
                            double currentValue = command.Parameters.ContainsKey("value") ? 
                                Convert.ToDouble(command.Parameters["value"]) : Math.Min(5.0, (minValue + maxValue) / 2);
                            
                            // 確保當前值在範圍內
                            currentValue = Math.Max(minValue, Math.Min(maxValue, currentValue));
                            
                            // 正確設置slider的屬性
                            slider.Slider.Minimum = (decimal)minValue;
                            slider.Slider.Maximum = (decimal)maxValue;
                            slider.Slider.DecimalPlaces = 2;
                            slider.SetSliderValue((decimal)currentValue);
                            
                            RhinoApp.WriteLine($"Created Number Slider: min={minValue}, max={maxValue}, value={currentValue}");
                            component = slider;
                            break;
                        case "number":
                        case "num":
                        case "param_number":
                            component = new Param_Number();
                            break;
                        case "integer":
                        case "int":
                        case "param_integer":
                            component = new Param_Integer();
                            break;
                        case "construct point":
                        case "constructpoint":
                        case "pt xyz":
                        case "xyz":
                            // 嘗試查找構造點組件
                            var pointProxy = Grasshopper.Instances.ComponentServer.ObjectProxies
                                .FirstOrDefault(p => p.Desc.Name.Equals("Construct Point", StringComparison.OrdinalIgnoreCase));
                            if (pointProxy != null)
                            {
                                component = pointProxy.CreateInstance();
                            }
                            else
                            {
                                throw new ArgumentException("Construct Point component not found");
                            }
                            break;
                        case "booleantoggle":
                        case "toggle":
                        case "bool":
                        case "boolean":
                            var toggle = new GH_BooleanToggle();
                            toggle.CreateAttributes();
                            
                            // 檢查是否有自定義值參數
                            bool boolValue = command.Parameters.ContainsKey("value") ? 
                                Convert.ToBoolean(command.Parameters["value"]) : false;
                            
                            toggle.Value = boolValue;
                            
                            RhinoApp.WriteLine($"Created Boolean Toggle: value={boolValue}");
                            component = toggle;
                            break;
                            
                        // 數學運算組件
                        case "addition":
                        case "add":
                        case "+":
                            component = CreateComponentByName("Addition");
                            break;
                        case "subtraction":
                        case "subtract":
                        case "-":
                            component = CreateComponentByName("Subtraction");
                            break;
                        case "multiplication":
                        case "multiply":
                        case "*":
                            component = CreateComponentByName("Multiplication");
                            break;
                        case "division":
                        case "divide":
                        case "/":
                            component = CreateComponentByName("Division");
                            break;
                            
                        // 列表組件
                        case "listitem":
                        case "item":
                            component = CreateComponentByName("List Item");
                            break;
                        case "listlength":
                        case "length":
                            component = CreateComponentByName("List Length");
                            break;
                        case "series":
                            component = CreateComponentByName("Series");
                            break;
                        case "range":
                            component = CreateComponentByName("Range");
                            break;
                            
                        // 變換組件
                        case "move":
                        case "translate":
                            component = CreateComponentByName("Move");
                            break;
                        case "rotate":
                            component = CreateComponentByName("Rotate");
                            break;
                        case "scale":
                            component = CreateComponentByName("Scale");
                            break;
                            
                        // 向量組件
                        case "vector2pt":
                        case "vec2pt":
                            component = CreateComponentByName("Vector 2Pt");
                            break;
                        case "distance":
                        case "dist":
                            component = CreateComponentByName("Distance");
                            break;
                            
                        // 曲線組件
                        case "loft":
                            component = CreateComponentByName("Loft");
                            break;
                        case "curvelength":
                            component = CreateComponentByName("Curve Length");
                            break;
                        case "evaluatecurve":
                        case "evalcrv":
                            component = CreateComponentByName("Evaluate Curve");
                            break;
                        case "dividecurve":
                        case "divcrv":
                            component = CreateComponentByName("Divide Curve");
                            break;
                        case "joincurves":
                        case "join":
                            component = CreateComponentByName("Join Curves");
                            break;
                        case "offsetcurve":
                        case "offset":
                            component = CreateComponentByName("Offset Curve");
                            break;
                        default:
                            // 嘗試通過 Guid 查找組件
                            Guid componentGuid;
                            if (Guid.TryParse(type, out componentGuid))
                            {
                                component = Grasshopper.Instances.ComponentServer.EmitObject(componentGuid);
                                RhinoApp.WriteLine($"Attempting to create component by GUID: {componentGuid}");
                            }
                            
                            if (component == null)
                            {
                                // 嘗試通過名稱查找組件（不區分大小寫）
                                RhinoApp.WriteLine($"Attempting to find component by name: {type}");
                                var obj = Grasshopper.Instances.ComponentServer.ObjectProxies
                                    .FirstOrDefault(p => p.Desc.Name.Equals(type, StringComparison.OrdinalIgnoreCase));
                                    
                                if (obj != null)
                                {
                                    RhinoApp.WriteLine($"Found component: {obj.Desc.Name}");
                                    component = obj.CreateInstance();
                                }
                                else
                                {
                                    // 嘗試通過部分名稱匹配
                                    RhinoApp.WriteLine($"Attempting to find component by partial name match: {type}");
                                    obj = Grasshopper.Instances.ComponentServer.ObjectProxies
                                        .FirstOrDefault(p => p.Desc.Name.IndexOf(type, StringComparison.OrdinalIgnoreCase) >= 0);
                                        
                                    if (obj != null)
                                    {
                                        RhinoApp.WriteLine($"Found component by partial match: {obj.Desc.Name}");
                                        component = obj.CreateInstance();
                                    }
                                }
                            }
                            
                            if (component == null)
                            {
                                // 記錄一些可能的組件類型
                                var possibleMatches = Grasshopper.Instances.ComponentServer.ObjectProxies
                                    .Where(p => p.Desc.Name.IndexOf(type, StringComparison.OrdinalIgnoreCase) >= 0)
                                    .Select(p => p.Desc.Name)
                                    .Take(10)
                                    .ToList();
                                
                                var errorMessage = $"Unknown component type: {type}";
                                if (possibleMatches.Any())
                                {
                                    errorMessage += $". Possible matches: {string.Join(", ", possibleMatches)}";
                                }
                                
                                throw new ArgumentException(errorMessage);
                            }
                            break;
                    }
                    
                    // 設置組件位置
                    if (component != null)
                    {
                        // 確保組件有有效的屬性對象
                        if (component.Attributes == null)
                        {
                            RhinoApp.WriteLine("Component attributes are null, creating new attributes");
                            component.CreateAttributes();
                        }
                        
                        // 設置位置
                        component.Attributes.Pivot = new System.Drawing.PointF((float)x, (float)y);
                        
                        // 添加到文檔
                        doc.AddObject(component, false);
                        
                        // 如果是 Slider，在添加到文檔後再設置值
                        if (component is GH_NumberSlider slider)
                        {
                            // 獲取參數
                            double minValue = command.Parameters.ContainsKey("min") ? Convert.ToDouble(command.Parameters["min"]) : 0.0;
                            double maxValue = command.Parameters.ContainsKey("max") ? Convert.ToDouble(command.Parameters["max"]) : 10.0;
                            double currentValue = command.Parameters.ContainsKey("value") ? Convert.ToDouble(command.Parameters["value"]) : (minValue + maxValue) / 2;

                            // 確保值在範圍內
                            currentValue = Math.Max(minValue, Math.Min(maxValue, currentValue));

                            // 延遲後設置，確保組件已初始化
                            RhinoApp.InvokeOnUiThread(new Action(() => {
                                slider.Slider.Minimum = (decimal)minValue;
                                slider.Slider.Maximum = (decimal)maxValue;
                                slider.SetSliderValue((decimal)currentValue);
                                slider.ExpireSolution(true); // 強制刷新
                                RhinoApp.WriteLine($"Slider '{slider.NickName}' value set to {currentValue}");
                            }));
                        }

                        // 刷新畫布
                        doc.NewSolution(false);
                        
                        // 返回組件信息
                        result = new
                        {
                            id = component.InstanceGuid.ToString(),
                            type = component.GetType().Name,
                            name = component.NickName,
                            x = component.Attributes.Pivot.X,
                            y = component.Attributes.Pivot.Y
                        };
                    }
                    else
                    {
                        throw new InvalidOperationException("Failed to create component");
                    }
                }
                catch (Exception ex)
                {
                    exception = ex;
                    RhinoApp.WriteLine($"Error in AddComponent: {ex.Message}");
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
                return Response.CreateError($"Error adding component: {exception.Message}");
            }

            return Response.Ok(new { result });
        }
        
        /// <summary>
        /// 連接組件
        /// </summary>
        /// <param name="command">包含源和目標組件信息的命令</param>
        /// <returns>連接信息</returns>
        public static object ConnectComponents(Command command)
        {
            var fromData = command.GetParameter<Dictionary<string, object>>("from");
            var toData = command.GetParameter<Dictionary<string, object>>("to");
            
            if (fromData == null || toData == null)
            {
                throw new ArgumentException("Source and target component information are required");
            }
            
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
                    
                    // 解析源組件信息
                    string fromIdStr = fromData["id"].ToString();
                    string fromParamName = fromData["parameterName"].ToString();
                    
                    // 解析目標組件信息
                    string toIdStr = toData["id"].ToString();
                    string toParamName = toData["parameterName"].ToString();
                    
                    // 將字符串 ID 轉換為 Guid
                    Guid fromId, toId;
                    if (!Guid.TryParse(fromIdStr, out fromId) || !Guid.TryParse(toIdStr, out toId))
                    {
                        throw new ArgumentException("Invalid component ID format");
                    }
                    
                    // 查找源和目標組件
                    IGH_Component fromComponent = doc.FindComponent(fromId) as IGH_Component;
                    IGH_Component toComponent = doc.FindComponent(toId) as IGH_Component;
                    
                    if (fromComponent == null || toComponent == null)
                    {
                        throw new ArgumentException("Source or target component not found");
                    }
                    
                    // 查找源輸出參數
                    IGH_Param fromParam = null;
                    foreach (var param in fromComponent.Params.Output)
                    {
                        if (param.Name.Equals(fromParamName, StringComparison.OrdinalIgnoreCase))
                        {
                            fromParam = param;
                            break;
                        }
                    }
                    
                    // 查找目標輸入參數
                    IGH_Param toParam = null;
                    foreach (var param in toComponent.Params.Input)
                    {
                        if (param.Name.Equals(toParamName, StringComparison.OrdinalIgnoreCase))
                        {
                            toParam = param;
                            break;
                        }
                    }
                    
                    if (fromParam == null || toParam == null)
                    {
                        throw new ArgumentException("Source or target parameter not found");
                    }
                    
                    // 連接參數
                    toParam.AddSource(fromParam);
                    
                    // 刷新畫布
                    doc.NewSolution(false);
                    
                    // 返回連接信息
                    result = new
                    {
                        from = new
                        {
                            id = fromComponent.InstanceGuid.ToString(),
                            name = fromComponent.NickName,
                            parameter = fromParam.Name
                        },
                        to = new
                        {
                            id = toComponent.InstanceGuid.ToString(),
                            name = toComponent.NickName,
                            parameter = toParam.Name
                        }
                    };
                }
                catch (Exception ex)
                {
                    exception = ex;
                    RhinoApp.WriteLine($"Error in ConnectComponents: {ex.Message}");
                }
            }));
            
            // 等待 UI 線程操作完成
            while (result == null && exception == null)
            {
                Thread.Sleep(10);
            }
            
            // 如果有異常，返回錯誤響應
            if (exception != null)
            {
                return Response.CreateError($"Error connecting components: {exception.Message}");
            }
            
            return Response.Ok(new { result });
        }
        
        /// <summary>
        /// 設置組件值
        /// </summary>
        /// <param name="command">包含組件 ID 和值的命令</param>
        /// <returns>操作結果</returns>
        public static object SetComponentValue(Command command)
        {
            // 检查命令参数是否存在
            if (command.Parameters == null)
            {
                return Response.CreateError("Command parameters are required");
            }
            
            // 更安全的参数获取方式
            string idStr = null;
            string value = null;
            
            try
            {
                idStr = command.GetParameter<string>("id");
                value = command.GetParameter<string>("value");
            }
            catch (Exception ex)
            {
                return Response.CreateError($"Error getting parameters: {ex.Message}");
            }
            
            if (string.IsNullOrEmpty(idStr))
            {
                return Response.CreateError("Component ID is required. Please provide a valid component ID.");
            }
            
            if (value == null)
            {
                return Response.CreateError("Value parameter is required.");
            }
            
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
                    
                    // 將字符串 ID 轉換為 Guid
                    Guid id;
                    if (!Guid.TryParse(idStr, out id))
                    {
                        throw new ArgumentException("Invalid component ID format");
                    }
                    
                    // 查找組件
                    IGH_DocumentObject component = doc.FindObject(id, true);
                    if (component == null)
                    {
                        throw new ArgumentException($"Component with ID {idStr} not found");
                    }
                    
                    // 根據組件類型設置值
                    if (component is GH_Panel panel)
                    {
                        panel.UserText = value;
                    }
                    else if (component is GH_NumberSlider slider)
                    {
                        double doubleValue;
                        if (double.TryParse(value, out doubleValue))
                        {
                            slider.SetSliderValue((decimal)doubleValue);
                        }
                        else
                        {
                            throw new ArgumentException("Invalid slider value format");
                        }
                    }
                    else if (component is IGH_Component ghComponent)
                    {
                        // 嘗試設置第一個輸入參數的值
                        if (ghComponent.Params.Input.Count > 0)
                        {
                            var param = ghComponent.Params.Input[0];
                            if (param is Param_String stringParam)
                            {
                                stringParam.PersistentData.Clear();
                                stringParam.PersistentData.Append(new Grasshopper.Kernel.Types.GH_String(value));
                            }
                            else if (param is Param_Number numberParam)
                            {
                                double doubleValue;
                                if (double.TryParse(value, out doubleValue))
                                {
                                    numberParam.PersistentData.Clear();
                                    numberParam.PersistentData.Append(new Grasshopper.Kernel.Types.GH_Number(doubleValue));
                                }
                                else
                                {
                                    throw new ArgumentException("Invalid number value format");
                                }
                            }
                            else
                            {
                                throw new ArgumentException($"Cannot set value for parameter type {param.GetType().Name}");
                            }
                        }
                        else
                        {
                            throw new ArgumentException("Component has no input parameters");
                        }
                    }
                    else
                    {
                        throw new ArgumentException($"Cannot set value for component type {component.GetType().Name}");
                    }
                    
                    // 刷新畫布
                    doc.NewSolution(false);
                    
                    // 返回操作結果
                    result = new
                    {
                        id = component.InstanceGuid.ToString(),
                        type = component.GetType().Name,
                        value = value
                    };
                }
                catch (Exception ex)
                {
                    exception = ex;
                    RhinoApp.WriteLine($"Error in SetComponentValue: {ex.Message}");
                }
            }));
            
            // 等待 UI 線程操作完成
            while (result == null && exception == null)
            {
                Thread.Sleep(10);
            }
            
            // 如果有異常，返回錯誤響應
            if (exception != null)
            {
                return Response.CreateError($"Error setting component value: {exception.Message}");
            }

            return Response.Ok(new { result });
        }
        
        /// <summary>
        /// 獲取組件信息
        /// </summary>
        /// <param name="command">包含組件 ID 的命令</param>
        /// <returns>組件信息</returns>
        public static object GetComponentInfo(Command command)
        {
            // Check if the ID parameter exists and is not null/empty
            if (!command.Parameters.ContainsKey("id") || 
                command.Parameters["id"] == null || 
                string.IsNullOrEmpty(command.Parameters["id"].ToString()))
            {
                throw new ArgumentException("Component ID is required. Please provide a valid component ID.");
            }
            
            string idStr = command.Parameters["id"].ToString();
            
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
                    
                    // 將字符串 ID 轉換為 Guid
                    Guid id;
                    if (!Guid.TryParse(idStr, out id))
                    {
                        throw new ArgumentException("Invalid component ID format");
                    }
                    
                    // 查找組件
                    IGH_DocumentObject component = doc.FindObject(id, true);
                    if (component == null)
                    {
                        throw new ArgumentException($"Component with ID {idStr} not found");
                    }
                    
                    // 收集組件信息
                    var componentInfo = new Dictionary<string, object>
                    {
                        { "id", component.InstanceGuid.ToString() },
                        { "type", component.GetType().Name },
                        { "name", component.NickName },
                        { "description", component.Description }
                    };
                    
                    // 如果是 IGH_Component，收集輸入和輸出參數信息
                    if (component is IGH_Component ghComponent)
                    {
                        var inputs = new List<Dictionary<string, object>>();
                        foreach (var param in ghComponent.Params.Input)
                        {
                            inputs.Add(new Dictionary<string, object>
                            {
                                { "name", param.Name },
                                { "nickname", param.NickName },
                                { "description", param.Description },
                                { "type", param.GetType().Name },
                                { "dataType", param.TypeName }
                            });
                        }
                        componentInfo["inputs"] = inputs;
                        
                        var outputs = new List<Dictionary<string, object>>();
                        foreach (var param in ghComponent.Params.Output)
                        {
                            outputs.Add(new Dictionary<string, object>
                            {
                                { "name", param.Name },
                                { "nickname", param.NickName },
                                { "description", param.Description },
                                { "type", param.GetType().Name },
                                { "dataType", param.TypeName }
                            });
                        }
                        componentInfo["outputs"] = outputs;
                    }
                    
                    // 如果是 GH_Panel，獲取其文本值
                    if (component is GH_Panel panel)
                    {
                        componentInfo["value"] = panel.UserText;
                    }
                    
                    // 如果是 GH_NumberSlider，獲取其值和範圍
                    if (component is GH_NumberSlider slider)
                    {
                        componentInfo["value"] = (double)slider.CurrentValue;
                        componentInfo["minimum"] = (double)slider.Slider.Minimum;
                        componentInfo["maximum"] = (double)slider.Slider.Maximum;
                    }
                    
                    result = componentInfo;
                }
                catch (Exception ex)
                {
                    exception = ex;
                    RhinoApp.WriteLine($"Error in GetComponentInfo: {ex.Message}");
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
        
        private static IGH_DocumentObject CreateComponentByName(string name)
        {
            var obj = Grasshopper.Instances.ComponentServer.ObjectProxies
                .FirstOrDefault(p => p.Desc.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                
            if (obj != null)
            {
                return obj.CreateInstance();
            }
            else
            {
                throw new ArgumentException($"Component with name {name} not found");
            }
        }
        
        /// <summary>
        /// 搜索組件
        /// </summary>
        /// <param name="command">包含搜索條件的命令</param>
        /// <returns>匹配的組件列表</returns>
        public static object SearchComponents(Command command)
        {
            string query = command.GetParameter<string>("query");
            if (string.IsNullOrEmpty(query))
            {
                throw new ArgumentException("Search query is required");
            }
            
            object result = null;
            Exception exception = null;
            
            // 在 UI 線程上執行
            RhinoApp.InvokeOnUiThread(new Action(() =>
            {
                try
                {
                    var components = new List<Dictionary<string, object>>();
                    
                    // 搜索所有可用的組件
                    var proxies = Grasshopper.Instances.ComponentServer.ObjectProxies;
                    
                    foreach (var proxy in proxies)
                    {
                        // 檢查名稱、類別或描述是否匹配查詢
                        if (proxy.Desc.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0 ||
                            proxy.Desc.Category.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0 ||
                            proxy.Desc.SubCategory.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0 ||
                            proxy.Desc.Description.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            components.Add(new Dictionary<string, object>
                            {
                                { "name", proxy.Desc.Name },
                                { "category", proxy.Desc.Category },
                                { "subcategory", proxy.Desc.SubCategory },
                                { "description", proxy.Desc.Description },
                                { "guid", proxy.Guid.ToString() }
                            });
                        }
                    }
                    
                    // 使用模糊匹配進行額外搜索
                    var fuzzyMatches = FuzzyMatcher.SearchComponents(query);
                    foreach (var match in fuzzyMatches)
                    {
                        // 避免重複添加
                        if (!components.Any(c => c["name"].ToString().Equals(match, StringComparison.OrdinalIgnoreCase)))
                        {
                            var proxy = proxies.FirstOrDefault(p => p.Desc.Name.Equals(match, StringComparison.OrdinalIgnoreCase));
                            if (proxy != null)
                            {
                                components.Add(new Dictionary<string, object>
                                {
                                    { "name", proxy.Desc.Name },
                                    { "category", proxy.Desc.Category },
                                    { "subcategory", proxy.Desc.SubCategory },
                                    { "description", proxy.Desc.Description },
                                    { "guid", proxy.Guid.ToString() }
                                });
                            }
                        }
                    }
                    
                    result = components;
                }
                catch (Exception ex)
                {
                    exception = ex;
                }
            }));
            
            // 等待 UI 線程完成
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
        /// 獲取所有組件
        /// </summary>
        /// <param name="command">命令參數</param>
        /// <returns>所有組件的列表</returns>
        public static object GetAllComponents(Command command)
        {
            try
            {
                var doc = Grasshopper.Instances.ActiveCanvas?.Document;
                if (doc == null)
                {
                    return Response.CreateError("No active Grasshopper document");
                }

                var components = new List<object>();
                
                foreach (var obj in doc.Objects)
                {
                    if (obj is IGH_Component component)
                    {
                        var componentInfo = new
                        {
                            id = component.InstanceGuid.ToString(),
                            type = component.Name,
                            name = component.NickName,
                            x = component.Attributes?.Pivot.X ?? 0,
                            y = component.Attributes?.Pivot.Y ?? 0,
                            description = component.Description,
                            category = component.Category,
                            subcategory = component.SubCategory,
                            inputCount = component.Params?.Input?.Count ?? 0,
                            outputCount = component.Params?.Output?.Count ?? 0,
                            inputs = component.Params?.Input?.Select(p => new
                            {
                                name = p.Name,
                                nickname = p.NickName,
                                description = p.Description,
                                type = p.Type.Name,
                                optional = p.Optional,
                                hasData = p.VolatileDataCount > 0
                            }).ToList(),
                            outputs = component.Params?.Output?.Select(p => new
                            {
                                name = p.Name,
                                nickname = p.NickName,
                                description = p.Description,
                                type = p.Type.Name,
                                hasData = p.VolatileDataCount > 0
                            }).ToList()
                        };
                        
                        components.Add(componentInfo);
                    }
                }
                
                return Response.Ok(new 
                { 
                    success = true, 
                    result = components,
                    count = components.Count
                });
            }
            catch (Exception ex)
            {
                return Response.CreateError($"Error getting all components: {ex.Message}");
            }
        }
    }
}
