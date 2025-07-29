# MCP Server for Microsoft PowerPoint

## Introduction

This MCP (Model Context Protocol) server provides a set of tools to create, manipulate, and manage Microsoft PowerPoint presentations. It allows a Roo agent to interact with PowerPoint files programmatically, automating tasks such as slide creation, content insertion, and formatting.

The server is built in Python and uses the `python-pptx` library.

## Features

- **Create new presentations** and open existing ones.
- **Add and configure slides** with different layouts.
- **Populate placeholders** with text content.
- **Add various elements** to slides:
    - Text boxes
    - Bullet points
    - Images (from local files or base64 strings)
    - Tables (with data)
    - Shapes
    - Charts
- **Format content**: Customize fonts, colors, alignment, etc.
- **Save presentations** to a specified file path.

## Prerequisites and Installation

1.  **Python 3.x**: Ensure Python is installed on your system.
2.  **Dependencies**: Install the required Python libraries using pip.

    ```bash
    pip install python-pptx
    ```

## Starting the Server

To start the MCP server, run the `ppt_mcp_server.py` script from your terminal:

```bash
python ppt_mcp_server.py
```

The server will start and listen for MCP requests from the agent.

## Usage

Once the server is running and connected, you can use its tools via the `use_mcp_tool` command. Here are a few examples:

### Example 1: Create a simple presentation

```xml
<use_mcp_tool>
  <server_name>powerpoint</server_name>
  <tool_name>create_presentation</tool_name>
  <arguments>
    {
      "id": "my_first_presentation"
    }
  </arguments>
</use_mcp_tool>
<use_mcp_tool>
  <server_name>powerpoint</server_name>
  <tool_name>add_slide</tool_name>
  <arguments>
    {
      "presentation_id": "my_first_presentation",
      "title": "Hello, World!"
    }
  </arguments>
</use_mcp_tool>
<use_mcp_tool>
  <server_name>powerpoint</server_name>
  <tool_name>save_presentation</tool_name>
  <arguments>
    {
      "presentation_id": "my_first_presentation",
      "file_path": "HelloWorld.pptx"
    }
  </arguments>
</use_mcp_tool>
```

### Example 2: Add a slide with bullet points

```xml
<use_mcp_tool>
  <server_name>powerpoint</server_name>
  <tool_name>add_slide</tool_name>
  <arguments>
    {
      "layout_index": 5,
      "title": "Key Features"
    }
  </arguments>
</use_mcp_tool>
<use_mcp_tool>
  <server_name>powerpoint</server_name>
  <tool_name>add_bullet_points</tool_name>
  <arguments>
    {
      "slide_index": 1,
      "placeholder_idx": 1,
      "bullet_points": [
        "Feature A: Does amazing things",
        "Feature B: Solves complex problems",
        "Feature C: Easy to use"
      ]
    }
  </arguments>
</use_mcp_tool>
