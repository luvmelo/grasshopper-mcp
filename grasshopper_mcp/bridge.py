import socket
import json
import os
import sys
import traceback
from typing import Dict, Any, Optional, List

# 使用 MCP 服務器
from mcp.server.fastmcp import FastMCP

# 設置 Grasshopper MCP 連接參數
GRASSHOPPER_HOST = "localhost"
GRASSHOPPER_PORT = 8080  # 默認端口，可以根據需要修改

# 創建 MCP 服務器
server = FastMCP("Grasshopper Bridge")

def send_to_grasshopper(command_type: str, params: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """向 Grasshopper MCP 發送命令"""
    if params is None:
        params = {}
    
    # 創建命令
    command = {
        "type": command_type,
        "parameters": params
    }
    
    try:
        print(f"Sending command to Grasshopper: {command_type} with params: {params}", file=sys.stderr)
        
        # 連接到 Grasshopper MCP
        client = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        client.connect((GRASSHOPPER_HOST, GRASSHOPPER_PORT))
        
        # 發送命令
        command_json = json.dumps(command)
        client.sendall((command_json + "\n").encode("utf-8"))
        print(f"Command sent: {command_json}", file=sys.stderr)
        
        # 接收響應
        response_data = b""
        while True:
            chunk = client.recv(4096)
            if not chunk:
                break
            response_data += chunk
            if response_data.endswith(b"\n"):
                break
        
        # 處理可能的 BOM
        response_str = response_data.decode("utf-8-sig").strip()
        print(f"Response received: {response_str}", file=sys.stderr)
        
        # 解析 JSON 響應
        response = json.loads(response_str)
        client.close()
        return response
    except Exception as e:
        print(f"Error communicating with Grasshopper: {str(e)}", file=sys.stderr)
        traceback.print_exc(file=sys.stderr)
        return {
            "success": False,
            "error": f"Error communicating with Grasshopper: {str(e)}"
        }

# 註冊 MCP 工具
@server.tool("add_component")
def add_component(component_type: str, x: float, y: float, min_value: float = None, max_value: float = None, value: float = None):
    """
    Add a component to the Grasshopper canvas
    
    Args:
        component_type: Component type (point, curve, circle, line, panel, slider)
        x: X coordinate on the canvas
        y: Y coordinate on the canvas
        min_value: Minimum value for sliders (optional)
        max_value: Maximum value for sliders (optional)
        value: Current value for sliders (optional)
    
    Returns:
        Result of adding the component
    """
    # 處理常見的組件名稱混淆問題
    component_mapping = {
        # Number Slider 的各種可能輸入方式
        "number slider": "Number Slider",
        "numeric slider": "Number Slider",
        "num slider": "Number Slider",
        "slider": "Number Slider",  # 當只提到 slider 且上下文是數值時，預設為 Number Slider
        
        # 其他組件的標準化名稱
        "md slider": "MD Slider",
        "multidimensional slider": "MD Slider",
        "multi-dimensional slider": "MD Slider",
        "graph mapper": "Graph Mapper",
        
        # 參數組件
        "boolean": "Boolean Toggle",
        "bool": "Boolean Toggle",
        "toggle": "Boolean Toggle",
        "integer": "Integer",
        "int": "Integer",
        
        # 數學運算組件
        "add": "Addition",
        "addition": "Addition",
        "plus": "Addition",
        "sum": "Addition",
        "+": "Addition",
        "subtract": "Subtraction",
        "subtraction": "Subtraction",
        "minus": "Subtraction",
        "difference": "Subtraction",
        "-": "Subtraction",
        "multiply": "Multiplication",
        "multiplication": "Multiplication",
        "times": "Multiplication",
        "product": "Multiplication",
        "*": "Multiplication",
        "divide": "Division",
        "division": "Division",
        "/": "Division",
        
        # 集合操作組件
        "list item": "List Item",
        "item": "List Item",
        "list length": "List Length",
        "length": "List Length",
        "series": "Series",
        "range": "Range",
        
        # 變換組件
        "move": "Move",
        "translate": "Move",
        "rotate": "Rotate",
        "scale": "Scale",
        
        # 向量組件
        "construct point": "Construct Point",
        "point": "Construct Point",
        "pt": "Construct Point",
        "vector 2pt": "Vector 2Pt",
        "vector": "Vector 2Pt",
        "distance": "Distance",
        "dist": "Distance",
        
        # 曲線組件
        "line": "Line",
        "rectangle": "Rectangle",
        "rect": "Rectangle",
        "circle": "Circle",
        "loft": "Loft",
        "curve length": "Curve Length",
        "evaluate curve": "Evaluate Curve",
        "divide curve": "Divide Curve",
        "join curves": "Join Curves",
        "offset curve": "Offset Curve",
        
        # 幾何組件
        "box": "Box",
        "cube": "Box",
        "sphere": "Sphere",
        "cylinder": "Cylinder",
        "cone": "Cone",
        
        # 輸出組件
        "panel": "Panel",
        "text panel": "Panel",
        "output panel": "Panel",
        "display": "Panel"
    }
    
    # 檢查並修正組件類型
    normalized_type = component_type.lower()
    if normalized_type in component_mapping:
        component_type = component_mapping[normalized_type]
        print(f"Component type normalized from '{normalized_type}' to '{component_mapping[normalized_type]}'", file=sys.stderr)
    
    params = {
        "type": component_type,
        "x": x,
        "y": y
    }
    
    # 如果是滑桿類型，添加範圍參數
    if "slider" in component_type.lower():
        if min_value is not None:
            params["min"] = min_value
        if max_value is not None:
            params["max"] = max_value
        if value is not None:
            params["value"] = value
        
        # 智能設置範圍：如果指定了值但沒有範圍，自動調整範圍
        if value is not None and min_value is None and max_value is None:
            if value > 1.0:
                # 如果值大於1，設置一個合理的範圍
                params["min"] = 0.0
                params["max"] = max(10.0, value * 2)
                print(f"Auto-adjusted slider range for value {value}: min=0.0, max={params['max']}", file=sys.stderr)
            elif value < 0:
                # 如果值是負數，設置包含負數的範圍
                params["min"] = min(value * 2, -10.0)
                params["max"] = max(10.0, abs(value) * 2)
                print(f"Auto-adjusted slider range for negative value {value}: min={params['min']}, max={params['max']}", file=sys.stderr)
            else:
                # 對於0-1之間的值，設置至少到2的範圍
                params["min"] = 0.0
                params["max"] = max(2.0, value * 2)
                print(f"Auto-adjusted slider range for small value {value}: min=0.0, max={params['max']}", file=sys.stderr)
        elif min_value is None and max_value is None:
            # 如果沒有值也沒有範圍，設置默認範圍
            params["min"] = 0.0
            params["max"] = 10.0
            print(f"Set default slider range: min=0.0, max=10.0", file=sys.stderr)
        
        print(f"Creating slider with params: {params}", file=sys.stderr)
    
    return send_to_grasshopper("add_component", params)

@server.tool("create_slider_with_value") 
def create_slider_with_value(x: float, y: float, target_value: float, name: str = ""):
    """
    Create a Number Slider with automatically adjusted range based on the target value
    
    Args:
        x: X coordinate on the canvas
        y: Y coordinate on the canvas
        target_value: The value that the slider needs to support
        name: Optional name for the slider
    
    Returns:
        Result of creating the slider
    """
    # 計算合適的範圍 - 改進邏輯確保範圍足夠大
    if target_value <= 0:
        min_val = min(target_value * 2, -10.0) if target_value < 0 else -10.0
        max_val = max(10.0, abs(target_value) * 2)
    elif target_value <= 1.0:
        min_val = 0.0
        max_val = max(2.0, target_value * 2)  # 確保最大值至少為2
    else:
        # 對於大於1的值，設置更寬的範圍
        min_val = 0.0
        max_val = max(target_value * 2, target_value + 10.0)  # 確保範圍足夠寬
    
    # 設置合適的精度
    if abs(target_value) < 0.1:
        precision = 0.01
    elif abs(target_value) < 1.0:
        precision = 0.1
    else:
        precision = 0.5
    
    print(f"Creating slider for value {target_value} with range [{min_val}, {max_val}]", file=sys.stderr)
    
    # 創建滑桿參數
    params = {
        "type": "Number Slider",
        "x": x,
        "y": y,
        "min": min_val,
        "max": max_val,
        "value": target_value,
        "rounding": precision
    }
    
    if name:
        params["name"] = name
    
    # 發送創建滑桿的命令
    result = send_to_grasshopper("add_component", params)
    
    # 檢查結果並記錄
    if result and result.get("success"):
        print(f"Successfully created slider with value {target_value}, range [{min_val}, {max_val}]", file=sys.stderr)
    else:
        print(f"Failed to create slider: {result.get('error', 'Unknown error')}", file=sys.stderr)
    
    return result

@server.tool("clear_document")
def clear_document():
    """Clear the Grasshopper document"""
    return send_to_grasshopper("clear_document")

@server.tool("save_document")
def save_document(path: str):
    """
    Save the Grasshopper document
    
    Args:
        path: Save path
    
    Returns:
        Result of the save operation
    """
    params = {
        "path": path
    }
    
    return send_to_grasshopper("save_document", params)

@server.tool("load_document")
def load_document(path: str):
    """
    Load a Grasshopper document
    
    Args:
        path: Document path
    
    Returns:
        Result of the load operation
    """
    params = {
        "path": path
    }
    
    return send_to_grasshopper("load_document", params)

@server.tool("get_document_info")
def get_document_info():
    """Get information about the Grasshopper document"""
    return send_to_grasshopper("get_document_info")

@server.tool("connect_components")
def connect_components(source_id: str, target_id: str, source_param: str = None, target_param: str = None, source_param_index: int = None, target_param_index: int = None):
    """
    Connect two components in the Grasshopper canvas
    
    Args:
        source_id: ID of the source component (output)
        target_id: ID of the target component (input)
        source_param: Name of the source parameter (optional)
        target_param: Name of the target parameter (optional)
        source_param_index: Index of the source parameter (optional, used if source_param is not provided)
        target_param_index: Index of the target parameter (optional, used if target_param is not provided)
    
    Returns:
        Result of connecting the components
    """
    # 獲取目標組件的信息，檢查是否已有連接
    target_info = send_to_grasshopper("get_component_info", {"id": target_id})
    
    # 檢查組件類型，如果是需要多個輸入的組件（如 Addition, Subtraction 等），智能分配輸入
    if target_info and "result" in target_info and "type" in target_info["result"]:
        component_type = target_info["result"]["type"]
        
        # 獲取現有連接
        connections = send_to_grasshopper("get_connections")
        existing_connections = []
        
        if connections and "result" in connections:
            for conn in connections["result"]:
                if conn.get("targetId") == target_id:
                    existing_connections.append(conn)
        
        # 對於特定需要多個輸入的組件，自動選擇正確的輸入端口
        if component_type in ["Addition", "Subtraction", "Multiplication", "Division", "Math"]:
            # 如果沒有指定目標參數，且已有連接到第一個輸入，則自動連接到第二個輸入
            if target_param is None and target_param_index is None:
                # 檢查第一個輸入是否已被佔用
                first_input_occupied = False
                for conn in existing_connections:
                    if conn.get("targetParam") == "A" or conn.get("targetParamIndex") == 0:
                        first_input_occupied = True
                        break
                
                # 如果第一個輸入已被佔用，則連接到第二個輸入
                if first_input_occupied:
                    target_param = "B"  # 第二個輸入通常命名為 B
                else:
                    target_param = "A"  # 否則連接到第一個輸入
    
    params = {
        "sourceId": source_id,
        "targetId": target_id
    }
    
    if source_param is not None:
        params["sourceParam"] = source_param
    elif source_param_index is not None:
        params["sourceParamIndex"] = source_param_index
        
    if target_param is not None:
        params["targetParam"] = target_param
    elif target_param_index is not None:
        params["targetParamIndex"] = target_param_index
    
    return send_to_grasshopper("connect_components", params)

@server.tool("create_pattern")
def create_pattern(description: str):
    """
    Create a pattern of components based on a high-level description
    
    Args:
        description: High-level description of what to create (e.g., '3D voronoi cube')
    
    Returns:
        Result of creating the pattern
    """
    params = {
        "description": description
    }
    
    return send_to_grasshopper("create_pattern", params)

@server.tool("get_available_patterns")
def get_available_patterns(query: str):
    """
    Get a list of available patterns that match a query
    
    Args:
        query: Query to search for patterns
    
    Returns:
        List of available patterns
    """
    params = {
        "query": query
    }
    
    return send_to_grasshopper("get_available_patterns", params)

@server.tool("get_component_info")
def get_component_info(component_id: str):
    """
    Get detailed information about a specific component
    
    Args:
        component_id: ID of the component to get information about
    
    Returns:
        Detailed information about the component, including inputs, outputs, and current values
    """
    params = {
        "id": component_id
    }
    
    result = send_to_grasshopper("get_component_info", params)
    
    # 增強返回結果，添加更多參數信息
    if result and "result" in result:
        component_data = result["result"]
        
        # 獲取組件類型
        if "type" in component_data:
            component_type = component_data["type"]
            
            # 查詢組件庫，獲取該類型組件的詳細參數信息
            component_library = get_component_library()
            if "components" in component_library:
                for lib_component in component_library["components"]:
                    if lib_component.get("name") == component_type or lib_component.get("fullName") == component_type:
                        # 將組件庫中的參數信息合併到返回結果中
                        if "settings" in lib_component:
                            component_data["availableSettings"] = lib_component["settings"]
                        if "inputs" in lib_component:
                            component_data["inputDetails"] = lib_component["inputs"]
                        if "outputs" in lib_component:
                            component_data["outputDetails"] = lib_component["outputs"]
                        if "usage_examples" in lib_component:
                            component_data["usageExamples"] = lib_component["usage_examples"]
                        if "common_issues" in lib_component:
                            component_data["commonIssues"] = lib_component["common_issues"]
                        break
            
            # 特殊處理某些組件類型
            if component_type == "Number Slider":
                # 嘗試從組件數據中獲取當前滑桿的實際設置
                if "currentSettings" not in component_data:
                    component_data["currentSettings"] = {
                        "min": component_data.get("min", 0),
                        "max": component_data.get("max", 10),
                        "value": component_data.get("value", 5),
                        "rounding": component_data.get("rounding", 0.1),
                        "type": component_data.get("type", "float")
                    }
            
            # 添加組件的連接信息
            connections = send_to_grasshopper("get_connections")
            if connections and "result" in connections:
                # 查找與該組件相關的所有連接
                related_connections = []
                for conn in connections["result"]:
                    if conn.get("sourceId") == component_id or conn.get("targetId") == component_id:
                        related_connections.append(conn)
                
                if related_connections:
                    component_data["connections"] = related_connections
    
    return result

@server.tool("get_all_components")
def get_all_components():
    """
    Get a list of all components in the current document
    
    Returns:
        List of all components in the document with their IDs, types, and positions
    """
    result = send_to_grasshopper("get_all_components")
    
    # 增強返回結果，為每個組件添加更多參數信息
    if result and "result" in result:
        components = result["result"]
        component_library = get_component_library()
        
        # 獲取所有連接信息
        connections = send_to_grasshopper("get_connections")
        connections_data = connections.get("result", []) if connections else []
        
        # 為每個組件添加詳細信息
        for component in components:
            if "id" in component and "type" in component:
                component_id = component["id"]
                component_type = component["type"]
                
                # 添加組件的詳細參數信息
                if "components" in component_library:
                    for lib_component in component_library["components"]:
                        if lib_component.get("name") == component_type or lib_component.get("fullName") == component_type:
                            # 將組件庫中的參數信息合併到組件數據中
                            if "settings" in lib_component:
                                component["availableSettings"] = lib_component["settings"]
                            if "inputs" in lib_component:
                                component["inputDetails"] = lib_component["inputs"]
                            if "outputs" in lib_component:
                                component["outputDetails"] = lib_component["outputs"]
                            break
                
                # 添加組件的連接信息
                related_connections = []
                for conn in connections_data:
                    if conn.get("sourceId") == component_id or conn.get("targetId") == component_id:
                        related_connections.append(conn)
                
                if related_connections:
                    component["connections"] = related_connections
                
                # 特殊處理某些組件類型
                if component_type == "Number Slider":
                    # 嘗試獲取滑桿的當前設置
                    component_info = send_to_grasshopper("get_component_info", {"id": component_id})
                    if component_info and "result" in component_info:
                        info_data = component_info["result"]
                        component["currentSettings"] = {
                            "min": info_data.get("min", 0),
                            "max": info_data.get("max", 10),
                            "value": info_data.get("value", 5),
                            "rounding": info_data.get("rounding", 0.1)
                        }
    
    return result

@server.tool("get_connections")
def get_connections():
    """
    Get a list of all connections between components in the current document
    
    Returns:
        List of all connections between components
    """
    return send_to_grasshopper("get_connections")

@server.tool("search_components")
def search_components(query: str):
    """
    Search for components by name or category
    
    Args:
        query: Search query
    
    Returns:
        List of components matching the search query
    """
    params = {
        "query": query
    }
    
    return send_to_grasshopper("search_components", params)

@server.tool("get_component_parameters")
def get_component_parameters(component_type: str):
    """
    Get a list of parameters for a specific component type
    
    Args:
        component_type: Type of component to get parameters for
    
    Returns:
        List of input and output parameters for the component type
    """
    params = {
        "componentType": component_type
    }
    
    return send_to_grasshopper("get_component_parameters", params)

@server.tool("validate_connection")
def validate_connection(source_id: str, target_id: str, source_param: str = None, target_param: str = None):
    """
    Validate if a connection between two components is possible
    
    Args:
        source_id: ID of the source component (output)
        target_id: ID of the target component (input)
        source_param: Name of the source parameter (optional)
        target_param: Name of the target parameter (optional)
    
    Returns:
        Whether the connection is valid and any potential issues
    """
    params = {
        "sourceId": source_id,
        "targetId": target_id
    }
    
    if source_param is not None:
        params["sourceParam"] = source_param
        
    if target_param is not None:
        params["targetParam"] = target_param
    
    return send_to_grasshopper("validate_connection", params)

@server.tool("set_component_value")
def set_component_value(component_id: str, value: str):
    """
    Set the value of a component (like a panel or slider)
    
    Args:
        component_id: The ID of the component to modify
        value: The new value to set
    
    Returns:
        Result of setting the component value
    """
    return send_to_grasshopper("set_component_value", {
        "id": component_id,
        "value": value
    })

@server.tool("execute_code")
def execute_code(code: str, language: str = "python", x: float = 0, y: float = 0, inputs: Optional[List[Dict[str, Any]]] = None, outputs: Optional[List[Dict[str, Any]]] = None):
    """
    Execute code in a Grasshopper Script component with parameter guidance
    
    Creates a Script component that executes Python or C# code. Script components in Grasshopper
    require manual parameter configuration through the UI.
    
    Args:
        code: The code to execute
        language: Programming language ("python" or "csharp")
        x: X position for the component (default: 0)
        y: Y position for the component (default: 0)
        inputs: List of intended input parameters for guidance. Each input should be a dict with:
            - name: Parameter name (e.g., "radius", "height") - use this as variable name in code
            - type: Parameter type ("number", "point", "curve", "brep", "string", "boolean", etc.)
            - description: Parameter description
        outputs: List of intended output parameters for guidance. Each output should be a dict with:
            - name: Parameter name (e.g., "geometry", "result") - use this as variable name in code
            - type: Parameter type ("number", "point", "curve", "brep", "string", "boolean", etc.)
            - description: Parameter description
    
    Manual Setup Required:
        1. After running this command, zoom in on the created Script component
        2. Use the ⊕ button to add input/output parameters
        3. Name parameters to match the variable names in your code
        4. Right-click on parameters to set type hints (Number, Point, Curve, etc.)
        5. Connect other components to the inputs for data flow
    
    Example:
        execute_code(
            code="import rhinoscriptsyntax as rs\na = rs.AddSphere([0,0,0], radius)",
            language="python",
            inputs=[{
                "name": "radius",
                "type": "number", 
                "description": "Sphere radius"
            }],
            outputs=[{
                "name": "sphere",
                "type": "brep",
                "description": "Generated sphere"
            }]
        )
        
        After running: manually add 'radius' input parameter and set type hint to 'Number'
    
    Returns:
        Result of creating the script component with code and parameter guidance
    """
    params = {
        "code": code,
        "language": language,
        "x": x,
        "y": y
    }
    
    if inputs:
        params["inputs"] = inputs
    
    if outputs:
        params["outputs"] = outputs
    
    return send_to_grasshopper("execute_code", params)

@server.tool("search_components")
def search_components(query: str):
    """
    Search for available Grasshopper components
    
    Args:
        query: Search query for component names, categories, or descriptions
    
    Returns:
        List of matching components with their information
    """
    return send_to_grasshopper("search_components", {
        "query": query
    })

@server.tool("get_all_components")
def get_all_components():
    """
    Get all available Grasshopper components
    
    Returns:
        List of all available components organized by category
    """
    return send_to_grasshopper("get_all_components", {})

@server.tool("generate_manual_setup_guide")
def generate_manual_setup_guide(
    script_component_id: str,
    code: str,
    language: str = "python",
    input_params: list = None,
    output_params: list = None
):
    """
    Generate a manual setup guide and display it in a Panel next to the script component.
    
    Args:
        script_component_id: The ID of the script component to place the guide next to.
        code: The script code to analyze.
        language: Programming language (python or csharp).
        input_params: List of expected input parameters.
        output_params: List of expected output parameters.
    
    Returns:
        Result of creating and setting the guide panel.
    """
    # 1. Get the script component's position
    try:
        script_info = send_to_grasshopper("get_component_info", {"id": script_component_id})
        if not script_info or not script_info.get("success"):
            return {
                "success": False,
                "error": "Failed to get script component info."
            }
        
        script_pos = script_info.get("data", {}).get("result", {})
        x = script_pos.get("x", 100)
        y = script_pos.get("y", 100)
    except Exception as e:
        return {
            "success": False,
            "error": f"Error getting component position: {e}"
        }

    # 2. Generate the guide text (reusing existing logic)
    guide_text = _generate_guide_text(code, language, input_params, output_params)

    # 3. Create a Panel to display the guide
    panel_x = x + 300  # Position the panel to the right
    panel_y = y
    
    try:
        add_panel_command = {
            "type": "Panel",
            "x": panel_x,
            "y": panel_y,
            "name": f"Setup Guide",
            "value": guide_text  # Pass the value directly
        }
        panel_info = send_to_grasshopper("add_component", add_panel_command)
        
        if not panel_info or not panel_info.get("success"):
            return {
                "success": False,
                "error": "Failed to create the guide panel."
            }
            
        panel_id = panel_info.get("data", {}).get("result", {}).get("id")

        # 4. Set the panel's content
        set_value_command = {
            "id": panel_id,
            "value": guide_text
        }
        
        set_value_result = send_to_grasshopper("set_component_value", set_value_command)
        
        # Optional: Connect the script component to the panel for clarity
        send_to_grasshopper("connect_components", {
            "sourceId": script_component_id,
            "targetId": panel_id
        })
        
        return set_value_result

    except Exception as e:
        return {
            "success": False,
            "error": f"Error creating guide panel: {e}"
        }

def _generate_guide_text(code, language, input_params, output_params):
    """Helper function to generate the markdown guide text."""
    if input_params is None:
        input_params = []
    if output_params is None:
        output_params = []
    
    # Analyze code for variables
    variables_in_code = analyze_code_variables(code, language)
    
    # Build the guide string
    guide = f"""SCRIPT SETUP GUIDE ({language.upper()})
1. Zoom in on script component to see ⊕/⊖
2. Use ⊕ to add/remove parameters.
3. Right-click params to set Name & Type Hint.

---INPUTS---
"""
    # ... (rest of the text generation logic)
    # This part is simplified for brevity but reuses the detailed logic from before
    
    all_inputs = {param.get("name"): param for param in input_params if isinstance(param, dict)}
    for var in variables_in_code.get("inputs", []):
        if var not in all_inputs:
            all_inputs[var] = {"type": "number", "description": f"Variable used in code: {var}"}

    if not all_inputs:
        guide += "None required.\n"
    else:
        for name, info in all_inputs.items():
            guide += f"- Name: {name}, Type: {info.get('type', 'number')}\n"

    guide += "\n---OUTPUTS---\n"
    all_outputs = {param.get("name"): param for param in output_params if isinstance(param, dict)}
    for var in variables_in_code.get("outputs", []):
        if var not in all_outputs:
            all_outputs[var] = {"type": "geometry", "description": f"Result variable: {var}"}

    if not all_outputs:
        guide += "None required.\n"
    else:
        for name, info in all_outputs.items():
            guide += f"- Name: {name}, Type: {info.get('type', 'geometry')}\n"
            
    return guide

def analyze_code_variables(code: str, language: str) -> dict:
    """
    Analyze code to find input/output variables and potential issues
    
    Args:
        code: The code to analyze
        language: Programming language
    
    Returns:
        Dictionary containing found variables and warnings
    """
    import re
    
    inputs = set()
    outputs = set()
    warnings = []
    
    lines = code.split('\n')
    
    if language.lower() == "python":
        # Python 变量分析
        for line in lines:
            line = line.strip()
            if not line or line.startswith('#'):
                continue
                
            # 查找赋值语句 (outputs)
            assignment_match = re.match(r'^(\w+)\s*=', line)
            if assignment_match:
                var_name = assignment_match.group(1)
                # 常见的输出变量模式
                if var_name.lower() in ['result', 'output', 'a', 'b', 'c', 'out', 'geometry', 'points', 'curves', 'surfaces']:
                    outputs.add(var_name)
            
            # 查找使用的变量 (potential inputs)
            # 简单的变量使用检测
            used_vars = re.findall(r'\b([a-zA-Z_][a-zA-Z0-9_]*)\b', line)
            for var in used_vars:
                # 跳过关键字、函数名、模块名等
                if var not in ['print', 'len', 'range', 'for', 'if', 'else', 'elif', 'while', 'def', 'class', 'import', 'from', 'as', 'True', 'False', 'None', 'and', 'or', 'not', 'in', 'is', 'math', 'ghpythonlib', 'Rhino', 'Grasshopper', 'System']:
                    if not assignment_match or var != assignment_match.group(1):
                        inputs.add(var)
    
    elif language.lower() == "csharp":
        # C# 变量分析
        for line in lines:
            line = line.strip()
            if not line or line.startswith('//'):
                continue
                
            # 查找赋值语句
            assignment_match = re.match(r'^(\w+)\s*=', line)
            if assignment_match:
                var_name = assignment_match.group(1)
                if var_name.lower() in ['result', 'output', 'a', 'b', 'c', 'out']:
                    outputs.add(var_name)
            
            # 查找使用的变量
            used_vars = re.findall(r'\b([a-zA-Z_][a-zA-Z0-9_]*)\b', line)
            for var in used_vars:
                if var not in ['var', 'int', 'double', 'string', 'bool', 'if', 'else', 'for', 'while', 'foreach', 'using', 'namespace', 'class', 'public', 'private', 'static', 'void', 'return', 'true', 'false', 'null', 'new', 'this', 'base', 'Point3d', 'Vector3d', 'Plane', 'Circle', 'Line', 'Curve', 'Surface', 'Brep', 'Mesh', 'Math', 'System', 'Rhino', 'Grasshopper']:
                    if not assignment_match or var != assignment_match.group(1):
                        inputs.add(var)
    
    # 移除明显的输出变量从输入列表
    inputs = inputs - outputs
    
    # 检查常见问题
    if not inputs and not outputs:
        warnings.append("No variables detected - code might be self-contained")
    
    return {
        "inputs": list(inputs),
        "outputs": list(outputs),
        "warnings": warnings
    }

# 註冊 MCP 資源
@server.resource("grasshopper://status")
def get_grasshopper_status():
    """Get Grasshopper status"""
    try:
        # 獲取文檔信息
        doc_info = send_to_grasshopper("get_document_info")
        
        # 獲取所有組件（使用增強版的 get_all_components）
        components_result = get_all_components()
        components = components_result.get("result", []) if components_result else []
        
        # 獲取所有連接
        connections = send_to_grasshopper("get_connections")
        
        # 添加常用組件的提示信息
        component_hints = {
            "Number Slider": {
                "description": "Single numeric value slider with adjustable range",
                "common_usage": "Use for single numeric inputs like radius, height, count, etc.",
                "parameters": ["min", "max", "value", "rounding", "type"],
                "creation_methods": {
                    "add_component": "add_component('Number Slider', x, y, min_value, max_value, value)",
                    "smart_creation": "create_slider_with_value(x, y, target_value) - automatically sets appropriate range"
                },
                "important_note": "Always specify min/max values or use create_slider_with_value for values > 1.0",
                "NOT_TO_BE_CONFUSED_WITH": "MD Slider (which is for multi-dimensional values)"
            },
            "MD Slider": {
                "description": "Multi-dimensional slider for vector input",
                "common_usage": "Use for vector inputs, NOT for simple numeric values",
                "NOT_TO_BE_CONFUSED_WITH": "Number Slider (which is for single numeric values)"
            },
            "Panel": {
                "description": "Displays text or numeric data",
                "common_usage": "Use for displaying outputs and debugging"
            },
            "Addition": {
                "description": "Adds two or more numbers",
                "common_usage": "Connect two Number Sliders to inputs A and B",
                "parameters": ["A", "B"],
                "connection_tip": "First slider should connect to input A, second to input B"
            }
        }
        
        # 為每個組件添加當前參數值的摘要
        component_summaries = []
        for component in components:
            summary = {
                "id": component.get("id", ""),
                "type": component.get("type", ""),
                "position": {
                    "x": component.get("x", 0),
                    "y": component.get("y", 0)
                }
            }
            
            # 添加組件特定的參數信息
            if "currentSettings" in component:
                summary["settings"] = component["currentSettings"]
            elif component.get("type") == "Number Slider":
                # 嘗試從組件信息中提取滑桿設置
                summary["settings"] = {
                    "min": component.get("min", 0),
                    "max": component.get("max", 10),
                    "value": component.get("value", 5),
                    "rounding": component.get("rounding", 0.1)
                }
            
            # 添加連接信息摘要
            if "connections" in component:
                conn_summary = []
                for conn in component["connections"]:
                    if conn.get("sourceId") == component.get("id"):
                        conn_summary.append({
                            "type": "output",
                            "to": conn.get("targetId", ""),
                            "sourceParam": conn.get("sourceParam", ""),
                            "targetParam": conn.get("targetParam", "")
                        })
                    else:
                        conn_summary.append({
                            "type": "input",
                            "from": conn.get("sourceId", ""),
                            "sourceParam": conn.get("sourceParam", ""),
                            "targetParam": conn.get("targetParam", "")
                        })
                
                if conn_summary:
                    summary["connections"] = conn_summary
            
            component_summaries.append(summary)
        
        return {
            "status": "Connected to Grasshopper",
            "document": doc_info.get("result", {}),
            "components": component_summaries,
            "connections": connections.get("result", []),
            "component_hints": component_hints,
            "recommendations": [
                "When needing a simple numeric input control, ALWAYS use 'Number Slider', not MD Slider",
                "For vector inputs (like 3D points), use 'MD Slider' or 'Construct Point' with multiple Number Sliders",
                "Use 'Panel' to display outputs and debug values",
                "When connecting multiple sliders to Addition, first slider goes to input A, second to input B",
                "IMPORTANT: For slider values > 1.0, use create_slider_with_value(x, y, target_value) or specify min/max values",
                "For sphere radius = 3.0, use create_slider_with_value(x, y, 3.0) to auto-set range [0.0, 6.0]"
            ],
            "canvas_summary": f"Current canvas has {len(component_summaries)} components and {len(connections.get('result', []))} connections"
        }
    except Exception as e:
        print(f"Error getting Grasshopper status: {str(e)}", file=sys.stderr)
        traceback.print_exc(file=sys.stderr)
        return {
            "status": f"Error: {str(e)}",
            "document": {},
            "components": [],
            "connections": []
        }

@server.resource("grasshopper://component_guide")
def get_component_guide():
    """Get guide for Grasshopper components and connections"""
    return {
        "title": "Grasshopper Component Guide",
        "description": "Guide for creating and connecting Grasshopper components",
        "components": [
            {
                "name": "Point",
                "category": "Params",
                "description": "Creates a point at specific coordinates",
                "inputs": [
                    {"name": "X", "type": "Number"},
                    {"name": "Y", "type": "Number"},
                    {"name": "Z", "type": "Number"}
                ],
                "outputs": [
                    {"name": "Pt", "type": "Point"}
                ]
            },
            {
                "name": "Circle",
                "category": "Curve",
                "description": "Creates a circle",
                "inputs": [
                    {"name": "Plane", "type": "Plane", "description": "Base plane for the circle"},
                    {"name": "Radius", "type": "Number", "description": "Circle radius"}
                ],
                "outputs": [
                    {"name": "C", "type": "Circle"}
                ]
            },
            {
                "name": "XY Plane",
                "category": "Vector",
                "description": "Creates an XY plane at the world origin or at a specified point",
                "inputs": [
                    {"name": "Origin", "type": "Point", "description": "Origin point", "optional": True}
                ],
                "outputs": [
                    {"name": "Plane", "type": "Plane", "description": "XY plane"}
                ]
            },
            {
                "name": "Addition",
                "fullName": "Addition",
                "description": "Adds two or more numbers",
                "inputs": [
                    {"name": "A", "type": "Number", "description": "First input value"},
                    {"name": "B", "type": "Number", "description": "Second input value"}
                ],
                "outputs": [
                    {"name": "Result", "type": "Number", "description": "Sum of inputs"}
                ],
                "usage_examples": [
                    "Connect two Number Sliders to inputs A and B to add their values",
                    "Connect multiple values to add them all together"
                ],
                "common_issues": [
                    "When connecting multiple sliders, ensure they connect to different inputs (A and B)",
                    "The first slider should connect to input A, the second to input B"
                ]
            },
            {
                "name": "Number Slider",
                "fullName": "Number Slider",
                "description": "Creates a slider for numeric input with adjustable range and precision",
                "inputs": [],
                "outputs": [
                    {"name": "N", "type": "Number", "description": "Number output"}
                ],
                "settings": {
                    "min": {"description": "Minimum value of the slider", "default": 0},
                    "max": {"description": "Maximum value of the slider", "default": 10},
                    "value": {"description": "Current value of the slider", "default": 5},
                    "rounding": {"description": "Rounding precision (0.01, 0.1, 1, etc.)", "default": 0.1},
                    "type": {"description": "Slider type (integer, floating point)", "default": "float"},
                    "name": {"description": "Custom name for the slider", "default": ""}
                },
                "usage_examples": [
                    "Create a Number Slider with min=0, max=100, value=50",
                    "Create a Number Slider for radius with min=0.1, max=10, value=2.5, rounding=0.1"
                ],
                "common_issues": [
                    "Confusing with other slider types",
                    "Not setting appropriate min/max values for the intended use"
                ],
                "disambiguation": {
                    "similar_components": [
                        {
                            "name": "MD Slider",
                            "description": "Multi-dimensional slider for vector input, NOT for simple numeric values",
                            "how_to_distinguish": "Use Number Slider for single numeric values; use MD Slider only when you need multi-dimensional control"
                        },
                        {
                            "name": "Graph Mapper",
                            "description": "Maps values through a graph function, NOT a simple slider",
                            "how_to_distinguish": "Use Number Slider for direct numeric input; use Graph Mapper only for function-based mapping"
                        }
                    ],
                    "correct_usage": "When needing a simple numeric input control, ALWAYS use 'Number Slider', not MD Slider or other variants"
                }
            },
            {
                "name": "Panel",
                "fullName": "Panel",
                "description": "Displays text or numeric data",
                "inputs": [
                    {"name": "Input", "type": "Any"}
                ],
                "outputs": []
            },
            {
                "name": "Math",
                "fullName": "Mathematics",
                "description": "Performs mathematical operations",
                "inputs": [
                    {"name": "A", "type": "Number"},
                    {"name": "B", "type": "Number"}
                ],
                "outputs": [
                    {"name": "Result", "type": "Number"}
                ],
                "operations": ["Addition", "Subtraction", "Multiplication", "Division", "Power", "Modulo"]
            },
            {
                "name": "Construct Point",
                "fullName": "Construct Point",
                "description": "Constructs a point from X, Y, Z coordinates",
                "inputs": [
                    {"name": "X", "type": "Number"},
                    {"name": "Y", "type": "Number"},
                    {"name": "Z", "type": "Number"}
                ],
                "outputs": [
                    {"name": "Pt", "type": "Point"}
                ]
            },
            {
                "name": "Line",
                "fullName": "Line",
                "description": "Creates a line between two points",
                "inputs": [
                    {"name": "Start", "type": "Point"},
                    {"name": "End", "type": "Point"}
                ],
                "outputs": [
                    {"name": "L", "type": "Line"}
                ]
            },
            {
                "name": "Extrude",
                "fullName": "Extrude",
                "description": "Extrudes a curve to create a surface or a solid",
                "inputs": [
                    {"name": "Base", "type": "Curve"},
                    {"name": "Direction", "type": "Vector"},
                    {"name": "Height", "type": "Number"}
                ],
                "outputs": [
                    {"name": "Brep", "type": "Brep"}
                ]
            }
        ],
        "connectionRules": [
            {
                "from": "Number",
                "to": "Circle.Radius",
                "description": "Connect a number to the radius input of a circle"
            },
            {
                "from": "Point",
                "to": "Circle.Plane",
                "description": "Connect a point to the plane input of a circle (not recommended, use XY Plane instead)"
            },
            {
                "from": "XY Plane",
                "to": "Circle.Plane",
                "description": "Connect an XY Plane to the plane input of a circle (recommended)"
            },
            {
                "from": "Number",
                "to": "Math.A",
                "description": "Connect a number to the first input of a Math component"
            },
            {
                "from": "Number",
                "to": "Math.B",
                "description": "Connect a number to the second input of a Math component"
            },
            {
                "from": "Number",
                "to": "Construct Point.X",
                "description": "Connect a number to the X input of a Construct Point component"
            },
            {
                "from": "Number",
                "to": "Construct Point.Y",
                "description": "Connect a number to the Y input of a Construct Point component"
            },
            {
                "from": "Number",
                "to": "Construct Point.Z",
                "description": "Connect a number to the Z input of a Construct Point component"
            },
            {
                "from": "Point",
                "to": "Line.Start",
                "description": "Connect a point to the start input of a Line component"
            },
            {
                "from": "Point",
                "to": "Line.End",
                "description": "Connect a point to the end input of a Line component"
            },
            {
                "from": "Circle",
                "to": "Extrude.Base",
                "description": "Connect a circle to the base input of an Extrude component"
            },
            {
                "from": "Number",
                "to": "Extrude.Height",
                "description": "Connect a number to the height input of an Extrude component"
            }
        ],
        "commonIssues": [
            "Using Point component instead of XY Plane for inputs that require planes",
            "Not specifying parameter names when connecting components",
            "Using incorrect component names (e.g., 'addition' instead of 'Math' with Addition operation)",
            "Trying to connect incompatible data types",
            "Not providing all required inputs for a component",
            "Using incorrect parameter names (e.g., 'A' and 'B' for Math component instead of the actual parameter names)",
            "Not checking if a connection was successful before proceeding"
        ],
        "tips": [
            "Always use XY Plane component for plane inputs",
            "Specify parameter names when connecting components",
            "For Circle components, make sure to use the correct inputs (Plane and Radius)",
            "Test simple connections before creating complex geometry",
            "Avoid using components that require selection from Rhino",
            "Use get_component_info to check the actual parameter names of a component",
            "Use get_connections to verify if connections were established correctly",
            "Use search_components to find the correct component name before adding it",
            "Use validate_connection to check if a connection is possible before attempting it"
        ]
    }

@server.resource("grasshopper://component_library")
def get_component_library():
    """Get a comprehensive library of Grasshopper components"""
    # 這個資源提供了一個更全面的組件庫，包括常用組件的詳細信息
    return {
        "categories": [
            {
                "name": "Params",
                "components": [
                    {
                        "name": "Point",
                        "fullName": "Point Parameter",
                        "description": "Creates a point parameter",
                        "inputs": [
                            {"name": "X", "type": "Number", "description": "X coordinate"},
                            {"name": "Y", "type": "Number", "description": "Y coordinate"},
                            {"name": "Z", "type": "Number", "description": "Z coordinate"}
                        ],
                        "outputs": [
                            {"name": "Pt", "type": "Point", "description": "Point output"}
                        ]
                    },
                    {
                        "name": "Number Slider",
                        "fullName": "Number Slider",
                        "description": "Creates a slider for numeric input with adjustable range and precision",
                        "inputs": [],
                        "outputs": [
                            {"name": "N", "type": "Number", "description": "Number output"}
                        ],
                        "settings": {
                            "min": {"description": "Minimum value of the slider", "default": 0},
                            "max": {"description": "Maximum value of the slider", "default": 10},
                            "value": {"description": "Current value of the slider", "default": 5},
                            "rounding": {"description": "Rounding precision (0.01, 0.1, 1, etc.)", "default": 0.1},
                            "type": {"description": "Slider type (integer, floating point)", "default": "float"},
                            "name": {"description": "Custom name for the slider", "default": ""}
                        },
                        "usage_examples": [
                            "Create a Number Slider with min=0, max=100, value=50",
                            "Create a Number Slider for radius with min=0.1, max=10, value=2.5, rounding=0.1"
                        ],
                        "common_issues": [
                            "Confusing with other slider types",
                            "Not setting appropriate min/max values for the intended use"
                        ],
                        "disambiguation": {
                            "similar_components": [
                                {
                                    "name": "MD Slider",
                                    "description": "Multi-dimensional slider for vector input, NOT for simple numeric values",
                                    "how_to_distinguish": "Use Number Slider for single numeric values; use MD Slider only when you need multi-dimensional control"
                                },
                                {
                                    "name": "Graph Mapper",
                                    "description": "Maps values through a graph function, NOT a simple slider",
                                    "how_to_distinguish": "Use Number Slider for direct numeric input; use Graph Mapper only for function-based mapping"
                                }
                            ],
                            "correct_usage": "When needing a simple numeric input control, ALWAYS use 'Number Slider', not MD Slider or other variants"
                        }
                    },
                    {
                        "name": "Panel",
                        "fullName": "Panel",
                        "description": "Displays text or numeric data",
                        "inputs": [
                            {"name": "Input", "type": "Any", "description": "Any input data"}
                        ],
                        "outputs": []
                    }
                ]
            },
            {
                "name": "Maths",
                "components": [
                    {
                        "name": "Math",
                        "fullName": "Mathematics",
                        "description": "Performs mathematical operations",
                        "inputs": [
                            {"name": "A", "type": "Number", "description": "First number"},
                            {"name": "B", "type": "Number", "description": "Second number"}
                        ],
                        "outputs": [
                            {"name": "Result", "type": "Number", "description": "Result of the operation"}
                        ],
                        "operations": ["Addition", "Subtraction", "Multiplication", "Division", "Power", "Modulo"]
                    }
                ]
            },
            {
                "name": "Vector",
                "components": [
                    {
                        "name": "XY Plane",
                        "fullName": "XY Plane",
                        "description": "Creates an XY plane at the world origin or at a specified point",
                        "inputs": [
                            {"name": "Origin", "type": "Point", "description": "Origin point", "optional": True}
                        ],
                        "outputs": [
                            {"name": "Plane", "type": "Plane", "description": "XY plane"}
                        ]
                    },
                    {
                        "name": "Construct Point",
                        "fullName": "Construct Point",
                        "description": "Constructs a point from X, Y, Z coordinates",
                        "inputs": [
                            {"name": "X", "type": "Number", "description": "X coordinate"},
                            {"name": "Y", "type": "Number", "description": "Y coordinate"},
                            {"name": "Z", "type": "Number", "description": "Z coordinate"}
                        ],
                        "outputs": [
                            {"name": "Pt", "type": "Point", "description": "Constructed point"}
                        ]
                    }
                ]
            },
            {
                "name": "Curve",
                "components": [
                    {
                        "name": "Circle",
                        "fullName": "Circle",
                        "description": "Creates a circle",
                        "inputs": [
                            {"name": "Plane", "type": "Plane", "description": "Base plane for the circle"},
                            {"name": "Radius", "type": "Number", "description": "Circle radius"}
                        ],
                        "outputs": [
                            {"name": "C", "type": "Circle", "description": "Circle output"}
                        ]
                    },
                    {
                        "name": "Line",
                        "fullName": "Line",
                        "description": "Creates a line between two points",
                        "inputs": [
                            {"name": "Start", "type": "Point", "description": "Start point"},
                            {"name": "End", "type": "Point", "description": "End point"}
                        ],
                        "outputs": [
                            {"name": "L", "type": "Line", "description": "Line output"}
                        ]
                    }
                ]
            },
            {
                "name": "Surface",
                "components": [
                    {
                        "name": "Extrude",
                        "fullName": "Extrude",
                        "description": "Extrudes a curve to create a surface or a solid",
                        "inputs": [
                            {"name": "Base", "type": "Curve", "description": "Base curve to extrude"},
                            {"name": "Direction", "type": "Vector", "description": "Direction of extrusion", "optional": True},
                            {"name": "Height", "type": "Number", "description": "Height of extrusion"}
                        ],
                        "outputs": [
                            {"name": "Brep", "type": "Brep", "description": "Extruded brep"}
                        ]
                    }
                ]
            }
        ],
        "dataTypes": [
            {
                "name": "Number",
                "description": "A numeric value",
                "compatibleWith": ["Number", "Integer", "Double"]
            },
            {
                "name": "Point",
                "description": "A 3D point in space",
                "compatibleWith": ["Point3d", "Point"]
            },
            {
                "name": "Vector",
                "description": "A 3D vector",
                "compatibleWith": ["Vector3d", "Vector"]
            },
            {
                "name": "Plane",
                "description": "A plane in 3D space",
                "compatibleWith": ["Plane"]
            },
            {
                "name": "Circle",
                "description": "A circle curve",
                "compatibleWith": ["Circle", "Curve"]
            },
            {
                "name": "Line",
                "description": "A line segment",
                "compatibleWith": ["Line", "Curve"]
            },
            {
                "name": "Curve",
                "description": "A curve object",
                "compatibleWith": ["Curve", "Circle", "Line", "Arc", "Polyline"]
            },
            {
                "name": "Brep",
                "description": "A boundary representation object",
                "compatibleWith": ["Brep", "Surface", "Solid"]
            }
        ]
    }

def main():
    """Main entry point for the Grasshopper MCP Bridge Server"""
    try:
        # 啟動 MCP 服務器
        print("Starting Grasshopper MCP Bridge Server...", file=sys.stderr)
        print("Please add this MCP server to Claude Desktop", file=sys.stderr)
        server.run()
    except Exception as e:
        print(f"Error starting MCP server: {str(e)}", file=sys.stderr)
        traceback.print_exc(file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
