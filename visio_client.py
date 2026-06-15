import os
import sys
import win32com.client
from config import AppConfig

class VisioClient:
    def __init__(self):
        pass

    def get_visio_application(self):
        import pythoncom
        pythoncom.CoInitialize()
        try:
            visio = win32com.client.GetActiveObject("Visio.Application")
        except Exception:
            visio = win32com.client.Dispatch("Visio.Application")
        visio.Visible = True
        return visio

    def init_new_visio_document(self):
        visio = self.get_visio_application()
        doc = visio.Documents.Add("")
        self.set_page_size(420, 297)
        return doc

    def init_new_draw_document(self):
        return self.init_new_visio_document()

    def set_page_size(self, width_mm: int, height_mm: int):
        visio = self.get_visio_application()
        page = visio.ActivePage
        page.PageSheet.Cells("PageWidth").FormulaU = f"{width_mm} mm"
        page.PageSheet.Cells("PageHeight").FormulaU = f"{height_mm} mm"

    def get_active_document(self):
        visio = self.get_visio_application()
        try:
            doc = visio.ActiveDocument
            if doc is None:
                doc = visio.Documents.Add("")
            return doc
        except Exception:
            return visio.Documents.Add("")

    def draw_erd_entity(self, name: str, attributes: list[str], x: int, y: int, width: int = 4000, height: int = 3000):
        if x < 1000:
            x = x * 100
        if y < 1000:
            y = y * 100
        if width < 1000:
            width = width * 100
        if height < 1000:
            height = height * 100
        width = max(width, 3500)
        height = max(height, 2000)

        max_len = max(len(name) + 2, 18)
        for attr in attributes:
            max_len = max(max_len, len(attr) + 2)
        required_w_in = max_len * 0.085 + 0.3

        num_lines = 2 + len(attributes)
        required_h_in = num_lines * 0.18 + 0.3

        x_in = (x / 100) * 0.0393700787
        y_in = (y / 100) * 0.0393700787
        w_in = max((width / 100) * 0.0393700787, required_w_in)
        h_in = max((height / 100) * 0.0393700787, required_h_in)

        visio = self.get_visio_application()
        page = visio.ActivePage
        page_height_in = page.PageSheet.Cells("PageHeight").ResultIU

        y_top_in = page_height_in - y_in
        y_bottom_in = y_top_in - h_in

        shape = page.DrawRectangle(x_in, y_bottom_in, x_in + w_in, y_top_in)

        lines = [f"[{name}]", "------------------"]
        for attr in attributes:
            lines.append(f"• {attr}")
        shape.Text = "\n".join(lines)

        shape.CellsSRC(4, 0, 0).FormulaU = "0"
        shape.CellsU("VerticalAlign").FormulaU = "0"
        shape.CellsU("LeftMargin").FormulaU = "2 mm"
        shape.CellsU("TopMargin").FormulaU = "2 mm"
        shape.CellsU("RightMargin").FormulaU = "2 mm"
        shape.CellsU("BottomMargin").FormulaU = "2 mm"

        return shape

    def find_shape_by_entity_name(self, name: str):
        visio = self.get_visio_application()
        page = visio.ActivePage
        target = f"[{name.strip().lower()}]"
        for shape in page.Shapes:
            if shape.Text:
                lines = shape.Text.split("\n")
                if lines and lines[0].strip().lower() == target:
                    return shape
        return None

    def connect_entities(self, from_entity_name: str, to_entity_name: str, cardinality: str):
        visio = self.get_visio_application()
        page = visio.ActivePage
        from_shape = self.find_shape_by_entity_name(from_entity_name)
        to_shape = self.find_shape_by_entity_name(to_entity_name)
        if from_shape is None or to_shape is None:
            missing = []
            if from_shape is None:
                missing.append(from_entity_name)
            if to_shape is None:
                missing.append(to_entity_name)
            return f"Error: Start or End shape not found for: {', '.join(missing)}. Please draw them first."

        connector = page.Drop(visio.Application.ConnectorToolDataObject, 0, 0)
        connector.CellsU("BeginX").GlueTo(from_shape.CellsU("PinX"))
        connector.CellsU("EndX").GlueTo(to_shape.CellsU("PinX"))

        card = cardinality.strip().lower()
        if "many-to-many" in card or "many to many" in card or "m:n" in card or "m-n" in card or "*:*" in card:
            connector.CellsU("BeginArrow").FormulaU = "28"
            connector.CellsU("EndArrow").FormulaU = "28"
            connector.Text = "M:N"
        elif "one-to-many" in card or "one to many" in card or "1:n" in card or "1:m" in card or "1-n" in card:
            connector.CellsU("BeginArrow").FormulaU = "25"
            connector.CellsU("EndArrow").FormulaU = "28"
            connector.Text = "1:N"
        elif "zero-to-many" in card or "zero to many" in card or "0:n" in card or "0..n" in card or "0..*" in card:
            connector.CellsU("BeginArrow").FormulaU = "25"
            connector.CellsU("EndArrow").FormulaU = "29"
            connector.Text = "0:N"
        elif "many-to-one" in card or "many to one" in card or "n:1" in card or "m:1" in card or "n-1" in card:
            connector.CellsU("BeginArrow").FormulaU = "28"
            connector.CellsU("EndArrow").FormulaU = "25"
            connector.Text = "N:1"
        elif "one-to-one" in card or "one to one" in card or "1:1" in card or "1-1" in card or "1..1" in card:
            connector.CellsU("BeginArrow").FormulaU = "25"
            connector.CellsU("EndArrow").FormulaU = "25"
            connector.Text = "1:1"
        elif "zero-to-one" in card or "zero to one" in card or "0:1" in card or "0..1" in card:
            connector.CellsU("BeginArrow").FormulaU = "25"
            connector.CellsU("EndArrow").FormulaU = "30"
            connector.Text = "0:1"
        else:
            connector.Text = cardinality

        return connector

    def save_as_vsdx(self, output_filename: str):
        base, _ = os.path.splitext(output_filename)
        output_filename = base + ".vsdx"
        output_filename = os.path.abspath(output_filename)
        doc = self.get_active_document()
        doc.SaveAs(output_filename)

    def save_as_odg(self, output_filename: str):
        self.save_as_vsdx(output_filename)
