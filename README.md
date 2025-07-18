# Grasshopper MCP Bridge

A Model Context Protocol (MCP) server that enables direct interaction with Grasshopper from AI assistants like Claude. This bridge allows AI to read, analyze, and manipulate Grasshopper definitions programmatically.

## 🎯 Major Update: Official API Support

**McNeel officially released new Script component APIs in Rhino 8 Service Release 18 (March 2025)!** 

Grasshopper MCP now fully supports these official APIs, enabling complete programmatic control of Script components without manual configuration.

## Features

### Core Functionality
- **Document Management**: Open, analyze, and manipulate Grasshopper definitions
- **Component Operations**: Get component information, create components, and manage connections
- **Geometry Analysis**: Extract and analyze geometric data from Grasshopper components
- **Intent Recognition**: Natural language understanding for Grasshopper operations

### Enhanced Code Execution (🆕 Official API)
- **Programmatic Parameter Configuration**: Automatically configure input/output parameters using official Rhino 8 SR18 APIs
- **Dual API Support**: Uses **`Python3Component.Create()`** and **`CSharpComponent.Create()`** methods (new API) with automatic fallback to legacy methods
- **Smart Type Hints**: Automatic type hint configuration using **`ScriptVariableParam`** and **`TypeHints.Select()`** methods
- **Automatic Parameter Configuration**: Creates fully functional Script components with automatic parameter setup (Rhino 8 SR18+) or guidance fallback (older versions)
- **15+ Parameter Types**: Support for number, string, point, curve, brep, mesh, geometry, and more
- **Multi-Language Support**: Both Python and C# script components
- **Intelligent Error Handling**: Graceful fallback between API versions

### System Architecture
- **API Priority**: Rhino 8 SR18 official APIs → Legacy APIs (automatic detection)
- **Component Types**: Creates modern Scripts components (not legacy Python components)
- **Parameter Management**: Full programmatic control over parameter creation, configuration, and type hints
- **Backward Compatibility**: Works with older Rhino versions through intelligent API detection

## Installation

1. **Install the MCP server:**
   ```bash
   pip install grasshopper-mcp
   ```
   
2. **Install the Grasshopper plugin:**
   - Copy `releases/GH_MCP.gha` to your Grasshopper Components folder
   - Restart Rhino/Grasshopper

3. **Configure Claude Desktop:**
   
   Add to your `claude_desktop_config.json`:
   ```json
   {
     "mcpServers": {
       "grasshopper": {
         "command": "python",
         "args": ["-m", "grasshopper_mcp"],
         "env": {
           "GRASSHOPPER_MCP_LOG_LEVEL": "INFO"
         }
       }
     }
   }
   ```

## Usage Examples

### 🆕 Automatic Parameter Configuration

Create a parametric sphere with automatic parameter setup:

```python
import grasshopper_mcp

# Creates a fully configured Script component
result = grasshopper_mcp.execute_code(
    code="""
import rhinoscriptsyntax as rs
import Rhino.Geometry as rg

# Create sphere
center = rg.Point3d(0, 0, 0)
sphere = rg.Sphere(center, radius)
a = sphere.ToBrep()
""",
    language="python",
    x=100,
    y=100,
    inputs=[
        {
            "name": "radius",
            "type": "number",
            "description": "Sphere radius",
            "optional": False
        }
    ],
    outputs=[
        {
            "name": "a",
            "type": "brep",
            "description": "Sphere geometry"
        }
    ]
)

print(f"Created: {result['message']}")
# Output: "Script component created with programmatically configured parameters using Rhino 8 SR18 API"
```

### 🆕 Complex Parametric Architecture

```python
# Create parametric building structure
result = grasshopper_mcp.execute_code(
    code="""
import Rhino.Geometry as rg
import math

# Create building columns
columns = []
for i in range(count):
    angle = (i * 2 * math.pi) / count
    x = radius * math.cos(angle)
    y = radius * math.sin(angle)
    
    base_pt = rg.Point3d(x, y, 0)
    top_pt = rg.Point3d(x, y, height)
    column = rg.Line(base_pt, top_pt)
    columns.append(column)

# Create roof
roof_center = rg.Point3d(0, 0, height + 2)
roof = rg.Sphere(roof_center, radius + 3)

a = columns  # Building columns
b = roof.ToBrep()  # Building roof
""",
    language="python",
    x=200,
    y=200,
    inputs=[
        {"name": "count", "type": "integer", "description": "Number of columns"},
        {"name": "radius", "type": "number", "description": "Building radius"},
        {"name": "height", "type": "number", "description": "Column height"}
    ],
    outputs=[
        {"name": "a", "type": "line", "description": "Building columns"},
        {"name": "b", "type": "brep", "description": "Building roof"}
    ]
)
```

### 🆕 C# Script Support

```python
# Create C# script component with automatic configuration
result = grasshopper_mcp.execute_code(
    code="""
using Rhino;
using Rhino.Geometry;
using System;

// Create complex geometry
Point3d center = new Point3d(0, 0, 0);
Sphere sphere = new Sphere(center, radius);
Brep sphereBrep = sphere.ToBrep();

// Apply transformation
Transform scale = Transform.Scale(center, factor);
sphereBrep.Transform(scale);

a = sphereBrep;
""",
    language="csharp",
    x=300,
    y=300,
    inputs=[
        {"name": "radius", "type": "number", "description": "Sphere radius"},
        {"name": "factor", "type": "number", "description": "Scale factor"}
    ],
    outputs=[
        {"name": "a", "type": "brep", "description": "Transformed sphere"}
    ]
)
```

### Traditional Operations

```python
# Document operations
doc_info = grasshopper_mcp.get_document_info()
print(f"Document: {doc_info['name']}")

# Component operations
components = grasshopper_mcp.get_all_components()
for comp in components:
    print(f"Component: {comp['name']} ({comp['type']})")

# Intent recognition
result = grasshopper_mcp.recognize_intent("create a 3D sphere geometry")
print(f"Intent: {result['intent']}")
```

## 🆕 API Comparison

### Before: Manual Setup Required
```text
1. Create Script component
2. Manually zoom to see ZUI interface
3. Manually click ⊕ buttons to add parameters
4. Manually set parameter names and type hints
5. Manually connect sliders and other components
```

### After: Fully Automatic
```text
1. Call execute_code with parameter specifications
2. System automatically creates proper Scripts component
3. System automatically configures all parameters
4. System automatically sets type hints
5. User only needs to connect data sources
```

## Supported Parameter Types

The system supports 15+ parameter types with automatic type hint configuration:

| Type | Description | Example |
|------|-------------|---------|
| `number` | Double precision float | `{"name": "radius", "type": "number"}` |
| `integer` | Integer number | `{"name": "count", "type": "integer"}` |
| `string` | Text string | `{"name": "text", "type": "string"}` |
| `boolean` | True/false value | `{"name": "enabled", "type": "boolean"}` |
| `point` | 3D point | `{"name": "center", "type": "point"}` |
| `vector` | 3D vector | `{"name": "direction", "type": "vector"}` |
| `plane` | Plane definition | `{"name": "base", "type": "plane"}` |
| `line` | Line geometry | `{"name": "axis", "type": "line"}` |
| `curve` | Curve geometry | `{"name": "profile", "type": "curve"}` |
| `surface` | Surface geometry | `{"name": "surface", "type": "surface"}` |
| `brep` | Solid geometry | `{"name": "solid", "type": "brep"}` |
| `mesh` | Mesh geometry | `{"name": "mesh", "type": "mesh"}` |
| `geometry` | Generic geometry | `{"name": "geometry", "type": "geometry"}` |
| `color` | Color value | `{"name": "color", "type": "color"}` |
| `matrix` | Transformation matrix | `{"name": "transform", "type": "matrix"}` |

## System Requirements

- **Rhino 8 SR18+** (recommended for full API support)
- **Rhino 7** (legacy API support)
- **Python 3.8+**
- **Claude Desktop** or compatible MCP client

## API Architecture

### Script Parameter Configuration

**Automatic Configuration** (Rhino 8 SR18+):
- Full automatic parameter setup using official McNeel APIs
- Uses `ScriptVariableParam` and `SetParametersToScript()` methods
- Parameters are immediately available after component creation

**Legacy Configuration** (Older versions):
- Programmatic parameter creation using `RegisterInputParam()` 
- Manual parameter configuration guidance provided
- Requires user to verify parameter setup

### Priority System
1. **Rhino 8 SR18 Official APIs** (Primary)
   - `Python3Component.Create()` / `CSharpComponent.Create()`
   - `ScriptVariableParam` for flexible parameters
   - `TypeHints.Select()` for type configuration
   - `VariableParameterMaintenance()` for state management

2. **Legacy APIs** (Automatic Fallback)
   - `GH_ComponentParamServer` methods
   - Standard Grasshopper parameter types
   - Automatic type mapping

### Error Handling
- Graceful API detection and fallback
- Comprehensive error logging
- Informative user feedback

## Available Commands

### Core Commands
- `get_document_info()` - Get current document information
- `get_all_components()` - List all components in document
- `get_component_info(id)` - Get detailed component information
- `connect_components(source_id, target_id)` - Connect components

### Enhanced Commands
- `execute_code()` - Create parameterized script components with full automation
- `recognize_intent()` - Natural language processing for Grasshopper operations

## Configuration

### Environment Variables
- `GRASSHOPPER_MCP_LOG_LEVEL` - Set logging level (DEBUG, INFO, WARNING, ERROR)
- `GRASSHOPPER_MCP_TIMEOUT` - Set command timeout (default: 30 seconds)

### Logging
Detailed logging is available for troubleshooting:
```text
INFO: Using Rhino 8 SR18 API for component creation
INFO: Script component created with programmatically configured parameters
INFO: Successfully configured 2 input parameters and 1 output parameter
```

## Development

### Building from Source
```bash
git clone [repository-url]
cd grasshopper-mcp
pip install -e .
```

### Building GHA Plugin
```bash
cd GH_MCP
dotnet build
```

### Running Tests
```bash
pytest tests/
```

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

We welcome contributions! Please see our contributing guidelines and open an issue or pull request.

## Changelog

### Version 2.0.0 (Latest)
- ✅ **Official API Support**: Full integration with Rhino 8 SR18 APIs
- ✅ **Automatic Parameter Configuration**: No manual setup required
- ✅ **Smart API Detection**: Automatic fallback system
- ✅ **Enhanced Type System**: 15+ parameter types with automatic type hints
- ✅ **Modern Script Components**: Creates new Scripts components (not legacy Python components)
- ✅ **Dual Language Support**: Python and C# with identical functionality
- ✅ **Improved Error Handling**: Comprehensive error management and logging

### Version 1.0.0
- Basic MCP server functionality
- Manual parameter configuration
- Legacy API support only

## Support

For issues and questions:
- Check the [documentation](ENHANCED_CODE_EXECUTION_GUIDE.md)
- Open an issue on GitHub
- Review system logs for troubleshooting

---

**Note**: This implementation uses the latest official McNeel APIs as documented in the Rhino 8 SR18 release notes and McNeel Developer Forum discussions. The system automatically detects available APIs and provides the best possible experience for your Rhino version.


setup instruction:
cd GH_MCP
dotnet build --configuration Release