# Create PowerPoint Presentation with ERD from Text Prompt

This project uses LangChain's MultiServerMCPClient to connect to a custom MCP server (Microsoft Office/Visio MCP) and generate a PowerPoint presentation containing an ERD (Entity-Relationship Diagram) visualization.

## Prerequisites

- **Python 3.8+**
- **Microsoft Office (PowerPoint/Visio)** installed
- **MCP Server** running with Visio/PowerPoint tools

## Installation

1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd microsoft-office-mcp
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Start MCP Server

Run the MCP server that exposes Visio/PowerPoint tools:

```bash
python server.py
```

### Run the Agent

Execute the agent script with a prompt describing the desired ERD:

```bash
python agent.py
```

Modify the `main()` function in `agent.py` to customize the prompt:

```python
async def main():
    agent = ERDAgent()
    prompt = """
    Create a PowerPoint presentation with an ERD diagram for this scenario:
    
    [Your ERD description here]
    
    Key requirements:
    - Use Crow's Foot notation for relationships
    - Keep entities within slide boundaries
    - Save as vsdx file
    """
    await agent.run(prompt)
```

## How It Works

1. **`server.py`**: Implements a FastMCP server that exposes:
   - `init_new_visio_document`: Creates a new Visio/PowerPoint document
   - `draw_erd_entity`: Draws ERD entity shapes with specified coordinates
   - `connect_entities`: Creates relationship connectors between entities
   - `save_as_vsdx`: Saves the diagram as a .vsdx file

2. **`agent.py`**: A LangGraph agent that:
   - Connects to the MCP server using `MultiServerMCPClient`
   - Uses `create_react_agent` with Ollama (or another LLM) for reasoning
   - Executes the MCP tools sequentially to build the diagram
   - Streams the agent's thought process to the console

## Customization

- **LLM Model**: Change `OLLAMA_MODEL` in `.env` or `config.py`
- **MCP Server URL**: Update `MCP_SERVER_URL` in `.env`
- **Diagram Design**: Modify the prompt in `agent.py` to change:
  - Entities and attributes
  - Relationships and cardinality
  - Document layout and styling

## Example Prompt

```python
"""
Design a simple ERD diagram for a gym membership system.

Entities:
- MEMBERS (member_id, name, email, phone, address)
- CLASSES (class_id, name, capacity, instructor)
- MEMBERSHIPS (membership_id, type, start_date, end_date)

Relationships:
- A MEMBER can have one MEMBERSHIP (1:1)
- A MEMBER can attend multiple CLASSES (M:N)
- A CLASS has one instructor (1:1)

Requirements:
- Use Crow's Foot notation
- Keep the diagram clean and organized
- Save as vsdx file with appropriate filename
"""
