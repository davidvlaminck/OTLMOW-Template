import asyncio
from pathlib import Path
from typing import Optional

from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.worksheet.worksheet import Worksheet

from otlmow_modelbuilder.OSLOCollector import OSLOCollector
from otlmow_modelbuilder.SQLDataClasses.OSLOClass import OSLOClass

from otlmow_template.TemplateGeneratorBase import TemplateGeneratorBase


class InfotemplateGenerator(TemplateGeneratorBase):
    """Generates an info template from a subset database.
    
    The info template provides a structural overview of the data model including
    classes, their attributes, and relations. It is intended for documentation
    and information purposes, not for data delivery.
    """
    
    def __init__(self):
        super().__init__()
        self.relation_dict: dict[str, set[tuple[str, str]]] = {}
    
    def _build_relation_dict(self):
        """Build relation dictionary from OSLORelaties using OSLOCollector."""
        self.relation_dict = {}
        
        for relation in self.collector.relations:
            bron = relation.bron_uri.split("#")[-1]
            doel = relation.doel_uri.split("#")[-1]
            relatie = relation.objectUri.split("#")[-1]
            
            self.relation_dict.setdefault(bron, set()).add((relatie, doel))
            self.relation_dict.setdefault(doel, set()).add((relatie, bron))
    
    def _get_attributes_for_class(self, class_uri: str) -> list[tuple]:
        """Get attributes for a class from the database (including inherited)."""
        import sqlite3
        conn = sqlite3.connect(self.collector.path)
        cursor = conn.cursor()
        
        cursor.execute("""
            SELECT label_nl, kardinaliteit_min, kardinaliteit_max, type
            FROM OSLOAttributen
            WHERE class_uri = ?
            ORDER BY label_nl
        """, (class_uri,))
        
        result = cursor.fetchall()
        conn.close()
        return result
    
    def _apply_template_layout(self, ws: Worksheet, attribute_count: int):
        """Apply standard layout formatting to worksheet."""
        ws.column_dimensions["A"].width = 40
        ws.column_dimensions["B"].width = 25
        ws.column_dimensions["C"].width = 30
        ws.column_dimensions["D"].width = 70
        ws.column_dimensions["E"].width = 45
        ws.column_dimensions["F"].width = 30
        ws.row_dimensions[2].height = 50
        
        bold = Font(bold=True)
        medium_border = Border(
            left=Side(style="medium"),
            right=Side(style="medium"),
            top=Side(style="medium"),
            bottom=Side(style="medium")
        )
        dotted_border = Border(
            left=Side(style="dotted"),
            right=Side(style="dotted"),
            top=Side(style="dotted"),
            bottom=Side(style="dotted")
        )
        yellow_fill = PatternFill(
            fill_type="solid",
            start_color="F4F1DE",
            end_color="F4F1DE"
        )
        
        for cell in ["A1", "A4", "A13"]:
            ws[cell].font = bold
            ws[cell].border = medium_border
        
        for cell in ["A14", "B14", "C14", "D14", "E14", "F14"]:
            ws[cell].font = bold
        
        last_row = 14 + attribute_count
        for row in range(15, last_row + 1):
            for col in range(1, 7):
                cell = ws.cell(row=row, column=col)
                cell.fill = yellow_fill
                cell.border = dotted_border
    
    def _create_worksheet_for_class(self, wb: Workbook, cls: OSLOClass) -> Worksheet:
        """Create a worksheet for a single class."""
        object_name = cls.name
        description = cls.definition
        
        ws = wb.create_sheet(title=object_name[:31])
        
        # Fixed structure
        ws["A1"] = "BESCHRIJVING"
        if description:
            ws.merge_cells("A2:F2")
            ws["A2"] = description
            ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
        
        ws["A4"] = "RELATIES"
        
        ws["B5"] = "Verplicht y/n"
        ws["C5"] = "Wie levert aan?"
        ws["D5"] = "Type Relatie"
        ws["E5"] = "Relatie Met"
        ws["F5"] = "Opmerkingen"
        
        ws["A13"] = "EIGENSCHAPPEN"
        
        ws["A14"] = "Attribuutnaam"
        ws["B14"] = "Verplicht y/n"
        ws["C14"] = "Wie levert aan?"
        ws["D14"] = "Wat?"
        ws["E14"] = "Formaat/Opmaak/naamconventie"
        ws["F14"] = "Opmerkingen"
        
        # Relations
        relaties = sorted(self.relation_dict.get(object_name, []))
        max_rel_rows = 7
        
        if len(relaties) > max_rel_rows:
            extra_rows = len(relaties) - max_rel_rows
            ws.insert_rows(13, extra_rows)
        
        eigenschappen_row = 13 + max(0, len(relaties) - max_rel_rows)
        
        ws[f"A{eigenschappen_row}"] = "EIGENSCHAPPEN"
        ws[f"A{eigenschappen_row+1}"] = "Attribuutnaam"
        ws[f"B{eigenschappen_row+1}"] = "Verplicht y/n"
        ws[f"C{eigenschappen_row+1}"] = "Wie levert aan?"
        ws[f"D{eigenschappen_row+1}"] = "Wat?"
        ws[f"E{eigenschappen_row+1}"] = "Formaat/Opmaak/naamconventie"
        ws[f"F{eigenschappen_row+1}"] = "Opmerkingen"
        
        relation_row = 6
        for relatie_type, relatie_met in relaties:
            ws.cell(relation_row, 4, relatie_type)
            ws.cell(relation_row, 5, relatie_met)
            relation_row += 1
        
        # Attributes
        attributes = self._get_attributes_for_class(cls.objectUri)
        
        attribuut_start = eigenschappen_row + 2
        current_row = attribuut_start
        
        for attribuut, kard_min, kard_max, datatype in attributes:
            ws.cell(current_row, 1, attribuut)
            
            if kard_min == "1":
                ws.cell(current_row, 2, "y")
            else:
                ws.cell(current_row, 2, "n")
            
            if datatype:
                datatype = datatype.split("#")[-1]
                ws.cell(current_row, 5, datatype)
            
            current_row += 1
        
        attribute_count = current_row - attribuut_start
        
        self._apply_template_layout(ws, attribute_count)
        
        return ws
    
    def generate_infotemplate_from_subset(
            self,
            subset_path: Path,
            template_file_path: Path,
            model_directory: Path = None):
        """Generate an info template from a subset file.
        
        :param subset_path: Path to the subset file
        :param template_file_path: Path to where the template file should be created
        :param model_directory: Path to the model directory (optional)
        """
        self.collector = self._load_collector_from_subset_path(subset_path)
        self.classes = self.collector.classes
        self._build_relation_dict()
        
        concrete_classes = self._filter_concrete_classes()
        
        wb = Workbook()
        wb.remove(wb.active)
        
        for cls in concrete_classes:
            print(f"Verwerken: {cls.name}")
            self._create_worksheet_for_class(wb, cls)
        
        wb.save(template_file_path)
        print("")
        print("=" * 40)
        print("KLAAR")
        print(template_file_path)
        print("=" * 40)
    
    async def generate_infotemplate_from_subset_async(
            self,
            subset_path: Path,
            template_file_path: Path,
            model_directory: Path = None):
        """Async version of generate_infotemplate_from_subset."""
        self.collector = await self._load_collector_async(subset_path)
        self.classes = self.collector.classes
        self._build_relation_dict()
        
        concrete_classes = self._filter_concrete_classes()
        
        wb = Workbook()
        wb.remove(wb.active)
        
        for cls in concrete_classes:
            print(f"Verwerken: {cls.name}")
            self._create_worksheet_for_class(wb, cls)
            await asyncio.sleep(0)
        
        wb.save(template_file_path)
        print("")
        print("=" * 40)
        print("KLAAR")
        print(template_file_path)
        print("=" * 40)