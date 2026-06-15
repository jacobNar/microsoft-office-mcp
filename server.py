from fastmcp import FastMCP
from visio_client import VisioClient

mcp = FastMCP("microsoft_visio_erd")
client = VisioClient()

@mcp.tool()
def init_new_visio_document():
    client.init_new_visio_document()
    return "New Visio document initialized."

@mcp.tool()
def init_new_draw_document():
    client.init_new_draw_document()
    return "New Visio document initialized (compatibility alias)."

@mcp.tool()
def draw_erd_entity(name: str, attributes: list[str], x: int, y: int, width: int = 4000, height: int = 3000):
    client.draw_erd_entity(name, attributes, x, y, width, height)
    return f"ERD entity '{name}' drawn."

@mcp.tool()
def connect_entities(from_entity_name: str, to_entity_name: str, cardinality: str):
    res = client.connect_entities(from_entity_name, to_entity_name, cardinality)
    if isinstance(res, str) and res.startswith("Error"):
        return res
    return f"Connected '{from_entity_name}' to '{to_entity_name}' with cardinality '{cardinality}'."

@mcp.tool()
def save_as_vsdx(output_filename: str):
    client.save_as_vsdx(output_filename)
    return f"Saved to '{output_filename}' as vsdx."

@mcp.tool()
def save_as_odg(output_filename: str):
    client.save_as_odg(output_filename)
    return f"Saved to '{output_filename}' as vsdx (compatibility alias)."

@mcp.tool()
def set_page_size(width_mm: int, height_mm: int):
    client.set_page_size(width_mm, height_mm)
    return f"Page size set to {width_mm}mm x {height_mm}mm."

init_new_visio_document.__doc__ = "Initialize a new Microsoft Visio document."
init_new_draw_document.__doc__ = "Initialize a new Microsoft Visio document (compatibility alias)."
draw_erd_entity.__doc__ = "Draw an ERD entity table shape. Coordinates (x, y) and dimensions (width, height) are in 1/100ths of a mm (e.g. x=5000, y=8000, width=4000, height=3000). Values < 1000 will be auto-scaled by 100."
connect_entities.__doc__ = "Connect two entities using a connector shape. Supports Crow's Foot database notation. Cardinality strings are parsed to set line endings: 'one-to-one' (1:1), 'one-to-many' (1:N), 'zero-to-many' (0:N), 'many-to-one' (N:1), 'zero-to-one' (0:1), and 'many-to-many' (M:N)."
save_as_vsdx.__doc__ = "Save the drawing canvas as a native Microsoft Visio (.vsdx) document."
save_as_odg.__doc__ = "Save the drawing canvas as a native Microsoft Visio (.vsdx) document (compatibility alias)."
set_page_size.__doc__ = "Set the drawing page dimensions in millimeters (e.g., width_mm=420, height_mm=297 for A3 landscape)."

if __name__ == "__main__":
    mcp.run(transport="sse", host="0.0.0.0", port=8000)
