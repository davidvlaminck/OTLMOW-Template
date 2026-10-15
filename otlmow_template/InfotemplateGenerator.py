from __future__ import annotations

import asyncio
from pathlib import Path
from typing import ClassVar

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet
from otlmow_modelbuilder.SQLDataClasses.OSLOClass import OSLOClass

from otlmow_template.TemplateGeneratorBase import TemplateGeneratorBase


class InfotemplateGenerator(TemplateGeneratorBase):
    """Generates an info template from a subset database.
    
    The info template provides a structural overview of the data model including
    classes, their attributes, and relations. It is intended for documentation
    and information purposes, not for data delivery.
    """
    
    # Feature 3: Generic attribute URIs that get light green background
    GENERIC_ATTRIBUTE_URIS: ClassVar[set[str]] = {
        'https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#AIMObject.assetId',
        'https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#AIMObject.datumOprichtingObject',
        'https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#AIMObject.bestekPostNummer',
        'https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#AIMNaamObject.naam',
        'https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#AIMDBStatus.isActief',
        'https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#AIMToestand.toestand',
        'https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#AIMVersie.assetVersie',
        'https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#AIMObject.typeURI',
        'https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#AIMObject.theoretischeLevensduur',
        'https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#AIMObject.standaardBestekPostNummer',
        'https://wegenenverkeer.data.vlaanderen.be/ns/implementatieelement#AIMObject.notitie',
    }
    
    # Lighter green color for generic attributes
    LIGHT_GREEN_COLOR = "C8F7C8"
    
    def __init__(self):
        super().__init__()
        self.relation_dict: dict[str, list[tuple[str, str]]] = {}
    
    def _build_relation_dict(self):
        """Build relation dictionary from OSLORelaties using OSLOCollector.
        
        Only includes relations where both bron and doel are concrete (non-abstract) classes.
        For 'heeftBetrokkene' relations, creates two entries (Feature 2).
        Uses lists to preserve duplicates for heeftBetrokkene.
        Stores direction from database: Unspecified=<===>, Source -> Destination===>.
        """
        self.relation_dict = {}
        
        # Build a set of concrete class names for quick lookup
        concrete_class_names = {cls.name for cls in self._filter_concrete_classes()}
        
        # Use a set to track unique relations for deduplication (except heeftBetrokkene)
        seen_relations = set()
        
        for relation in self.collector.relations:
            bron = relation.bron_uri.split("#")[-1]
            doel = relation.doel_uri.split("#")[-1]
            relatie = relation.objectUri.split("#")[-1]
            richting = relation.richting  # Get direction from database
            
            # Only include relations where both bron and doel are concrete classes
            if bron in concrete_class_names and doel in concrete_class_names:
                # Add relation from bron perspective
                if relatie == 'heeftBetrokkene':
                    # For heeftBetrokkene, allow duplicates
                    self.relation_dict.setdefault(bron, []).append((relatie, doel, richting))
                else:
                    rel_key = (bron, relatie, doel)
                    if rel_key not in seen_relations:
                        seen_relations.add(rel_key)
                        self.relation_dict.setdefault(bron, []).append((relatie, doel, richting))
                
                # Also add from doel perspective (reverse direction)
                if relatie == 'heeftBetrokkene':
                    self.relation_dict.setdefault(doel, []).append((relatie, bron, richting))
                else:
                    rel_key = (doel, relatie, bron)
                    if rel_key not in seen_relations:
                        seen_relations.add(rel_key)
                        self.relation_dict.setdefault(doel, []).append((relatie, bron, richting))
        
        # Feature 2: Double the 'heeftBetrokkene' relations
        for bron, relations in self.relation_dict.items():
            new_relations = []
            for relatie, doel, richting in relations:
                new_relations.append((relatie, doel, richting))
                if relatie == 'heeftBetrokkene':
                    # Add a second entry for heeftBetrokkene
                    new_relations.append((relatie, doel, richting))
            self.relation_dict[bron] = new_relations
    
    def _get_attributes_for_class(self, class_uri: str) -> list[tuple]:
        """Get attributes for a class from the database (including inherited).

        Returns list of tuples: (label_nl, kardinaliteit_min, kardinaliteit_max, type, attribute_uri, definitie)
        """
        import sqlite3
        conn = sqlite3.connect(self.collector.path)
        cursor = conn.cursor()

        # Try to get definitie column; if it doesn't exist, default to None
        try:
            cursor.execute("""
                SELECT label_nl, kardinaliteit_min, kardinaliteit_max, type, uri, definition_nl
                FROM OSLOAttributen
                WHERE class_uri = ?
                ORDER BY label_nl
            """, (class_uri,))
        except sqlite3.OperationalError:
            try:
                cursor.execute("""
                    SELECT label_nl, kardinaliteit_min, kardinaliteit_max, type, uri, definitie
                    FROM OSLOAttributen
                    WHERE class_uri = ?
                    ORDER BY label_nl
                """, (class_uri,))
            except sqlite3.OperationalError:
                # Definitie column does not exist, select without it
                cursor.execute("""
                    SELECT label_nl, kardinaliteit_min, kardinaliteit_max, type, uri
                    FROM OSLOAttributen
                    WHERE class_uri = ?
                    ORDER BY label_nl
                """, (class_uri,))
                # Add a None for definitie for each row
                result = [row + (None,) for row in cursor.fetchall()]
                conn.close()
                return result

        result = cursor.fetchall()
        conn.close()
        return result

    def _get_codelist_uri_for_type(self, type_uri: str) -> str | None:
        """Get the codelist URI for a given enumeration type from OSLOEnumeration table.

        :param type_uri: The type URI (e.g. 'https://...#KlAIMToestand')
        :return: The codelist URI if found, None otherwise
        """
        if not type_uri:
            return None

        import sqlite3
        conn = sqlite3.connect(self.collector.path)
        cursor = conn.cursor()

        # Extract the type name from URI (e.g. 'KlAIMToestand' from '...#KlAIMToestand')
        type_name = type_uri.split('#')[-1] if '#' in type_uri else type_uri

        # Check if OSLOEnumeration table exists
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='OSLOEnumeration'")
        if not cursor.fetchone():
            conn.close()
            return None

        # Query the codelist for this enumeration
        cursor.execute("""
            SELECT codelist
            FROM OSLOEnumeration
            WHERE name = ?
        """, (type_name,))

        row = cursor.fetchone()
        conn.close()

        if row and row[0]:
            return row[0]
        return None
    
    def _apply_template_layout(self, ws: Worksheet, attribute_count: int, attr_header_row: int, rel_start_row: int, rel_end_row: int, eigenschappen_row: int):
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
        
        for cell in ["A1", "A4"]:
            ws[cell].font = bold
            ws[cell].border = medium_border

        # Set the "EIGENSCHAPPEN" header (column A of eigenschappen_row) to bold and medium border
        ws.cell(row=eigenschappen_row, column=1).font = bold
        ws.cell(row=eigenschappen_row, column=1).border = medium_border
        
        # Feature 6: Bold relation headers (row 5)
        for cell in ["A5", "B5", "C5", "D5", "E5", "F5"]:
            ws[cell].font = bold
        
        # Feature 6: Bold attribute headers
        for col in range(1, 7):
            cell = ws.cell(row=attr_header_row, column=col)
            cell.font = bold
        
        last_row = attr_header_row + attribute_count
        for row in range(attr_header_row + 1, last_row + 1):
            for col in range(1, 7):
                cell = ws.cell(row=row, column=col)
                cell.border = dotted_border
                # Only apply yellow fill if not already set (generic attributes get light green)
                if cell.fill.start_color.rgb == '00000000':
                    cell.fill = yellow_fill
        
        # Apply dotted border to relation rows too
        for row in range(rel_start_row, rel_end_row + 1):
            for col in range(1, 7):
                cell = ws.cell(row=row, column=col)
                cell.border = dotted_border
                if cell.fill.start_color.rgb == '00000000':
                    cell.fill = yellow_fill
        
        # Feature 10: Auto-filter on attribute header row
        ws.auto_filter.ref = f"A{attr_header_row}:F{last_row}"
        
        # Data validation for "Wie levert aan?" column (C) - both relations and attributes
        dv = DataValidation(type="list", formula1='"AWV,aannemer"', allow_blank=True)
        dv.error = "Selecteer AWV of aannemer"
        dv.errorTitle = "Ongeldige invoer"
        ws.add_data_validation(dv)
        
        # Apply to relation rows (column C)
        for row in range(rel_start_row, rel_end_row + 1):
            dv.add(ws.cell(row=row, column=3))
        
        # Apply to attribute rows (column C)
        for row in range(attr_header_row + 2, last_row + 1):
            dv.add(ws.cell(row=row, column=3))
        
        # Data validation for "Verplicht ja/nee" column (B) - attribute rows only
        dv_verplicht = DataValidation(type="list", formula1='"ja,nee"', allow_blank=True)
        dv_verplicht.error = "Selecteer ja of nee"
        dv_verplicht.errorTitle = "Ongeldige invoer"
        ws.add_data_validation(dv_verplicht)
        
        # Apply to attribute rows (column B)
        for row in range(attr_header_row + 1, last_row + 1):
            dv_verplicht.add(ws.cell(row=row, column=2))
    
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
        
        # Relations headers: A=Relatie Met, B=Verplicht y/n, C=Wie levert aan?, D=Type Relatie, E=Richting, F=Opmerkingen
        ws["A5"] = "Relatie Met"
        ws["B5"] = "Verplicht ja/nee"
        ws["C5"] = "Wie levert aan?"
        ws["D5"] = "Type Relatie"
        ws["E5"] = "Richting"
        ws["F5"] = "Opmerkingen"
        

        
        # Relations
        relaties = sorted(self.relation_dict.get(object_name, []))
        rel_start_row = 6
        rel_end_row = rel_start_row + len(relaties) - 1  # if len(relaties)==0, then rel_end_row=5

        # Write relations: A=Relatie Met, D=Type Relatie, E=Richting
        relation_row = rel_start_row
        for relatie_type, relatie_met, richting in relaties:
            ws.cell(relation_row, 1, relatie_met)      # A = Relatie Met
            ws.cell(relation_row, 4, relatie_type)      # D = Type Relatie
            # Feature 11: Column E is now "Richting" with direction indicators from database
            # Unspecified = ⇔ (undirected), Source -> Destination = ⇒ (directed from bron to doel)
            # Using Unicode arrows to avoid Excel/LibreOffice formula interpretation issues
            if richting == 'Source -> Destination':
                richting_symbol = "⇒"
            else:
                richting_symbol = "⇔"
            ws.cell(relation_row, 5, richting_symbol)   # E = Richting
            # Feature 4: Default 'aannemer' in "Wie levert aan?" column (C)
            ws.cell(relation_row, 3, 'aannemer')
            # B = Verplicht ja/nee (leave empty for relations)
            relation_row += 1

        # Reserve minimum space for relation table (like original code with max_rel_rows=7)
        # This ensures attributes don't start too early when there are few relations
        min_rel_end_row = 12  # rows 6-12 reserved for relation data (7 rows)
        relation_table_last_row = max(rel_end_row, min_rel_end_row)  # header is at row 5, data rows from 6 to relation_table_last_row
        # Two empty rows after the relation table
        eigenschappen_row = relation_table_last_row + 3
        attr_header_row = eigenschappen_row + 1
        attribuut_start = eigenschappen_row + 2

        # Set the "EIGENSCHAPPEN" header
        ws.cell(row=eigenschappen_row, column=1, value="EIGENSCHAPPEN")
        # Set attribute headers
        ws.cell(row=attr_header_row, column=1, value="Attribuutnaam")
        ws.cell(row=attr_header_row, column=2, value="Verplicht ja/nee")
        ws.cell(row=attr_header_row, column=3, value="Wie levert aan?")
        ws.cell(row=attr_header_row, column=4, value="Wat?")
        ws.cell(row=attr_header_row, column=5, value="Formaat/Opmaak/naamconventie")
        ws.cell(row=attr_header_row, column=6, value="Opmerkingen")
        
        # Feature 3: Light green fill for generic attributes
        light_green_fill = PatternFill(
            fill_type="solid",
            start_color=self.LIGHT_GREEN_COLOR,
            end_color=self.LIGHT_GREEN_COLOR
        )
        
        # Attributes - get all attributes first, then sort: generic first (alphabetically), then non-generic (alphabetically)
        attributes = self._get_attributes_for_class(cls.objectUri)
        
        # Sort: generic attributes first (by name), then non-generic (by name)
        generic_attrs = []
        non_generic_attrs = []
        for attr in attributes:
            attribuut, kard_min, kard_max, datatype, attr_uri, definitie = attr
            if attr_uri in self.GENERIC_ATTRIBUTE_URIS:
                generic_attrs.append(attr)
            else:
                non_generic_attrs.append(attr)
        
        generic_attrs.sort(key=lambda x: x[0].lower())  # Sort by attribute name
        non_generic_attrs.sort(key=lambda x: x[0].lower())
        
        sorted_attributes = generic_attrs + non_generic_attrs

        attribuut_start = eigenschappen_row + 2
        current_row = attribuut_start

        for attribuut, kard_min, kard_max, datatype, attr_uri, definitie in sorted_attributes:
            ws.cell(current_row, 1, attribuut)

            if kard_min == "1":
                ws.cell(current_row, 2, "ja")
            else:
                ws.cell(current_row, 2, "nee")

            # Write definition to "Wat?" column (D)
            ws.cell(current_row, 4, definitie if definitie is not None else "")

            # Feature 4: Default 'aannemer' in "Wie levert aan?" column (C)
            ws.cell(current_row, 3, 'aannemer')

            # Special opmerkingen for certain attributes
            attr_lower = attribuut.lower()
            if attr_lower == 'is actief':
                ws.cell(current_row, 6, "Moet TRUE zijn, behalve wanneer het asset uit de databank verwijderd moet worden")
            elif attr_lower == 'toestand' or attr_lower == 'aim-toestand':
                ws.cell(current_row, 6, "Is bijna altijd 'in-gebruik'")

            # Handle datatype in Formaat/Opmaak/naamconventie (col E)
            if datatype:
                datatype_clean = datatype.split("#")[-1]
                ws.cell(current_row, 5, datatype_clean)

                # Feature 8: If type starts with 'Kl' (keuzelijst/codelist), add codelist URI link
                if datatype_clean.startswith('Kl'):
                    codelist_uri = self._get_codelist_uri_for_type(datatype)
                    if codelist_uri:
                        ws.cell(current_row, 5, f"{datatype_clean} - {codelist_uri}")

            # Feature 5: TypeURI in opmerkingen for typeURI attribute only
            if attribuut.lower() == 'type uri':
                # Get the class typeURI
                class_typeuri = cls.objectUri
                ws.cell(current_row, 6, f"typeURI: {class_typeuri}")

            # Feature 3: Apply light green fill for generic attributes
            if attr_uri in self.GENERIC_ATTRIBUTE_URIS:
                for col in range(1, 7):
                    cell = ws.cell(current_row, col)
                    cell.fill = light_green_fill
                    cell.border = Border(
                        left=Side(style="dotted"),
                        right=Side(style="dotted"),
                        top=Side(style="dotted"),
                        bottom=Side(style="dotted")
                    )

            current_row += 1

        attribute_count = current_row - attribuut_start
        
        self._apply_template_layout(ws, attribute_count, attr_header_row, rel_start_row, rel_end_row, eigenschappen_row)
        
        return ws
    
    def generate_infotemplate_from_subset(
            self,
            subset_path: Path,
            template_file_path: Path,
            model_directory: Path | None = None):
        """Generate an info template from a subset file.
        
        :param subset_path: Path to the subset file
        :param template_file_path: Path to where the template file should be created
        :param model_directory: Path to the model directory (optional)
        """
        self.collector = self._load_collector_from_subset_path(subset_path)
        self.classes = self.collector.classes
        self._build_relation_dict()
        
        concrete_classes = self._filter_concrete_classes()
        # Feature 12: Sort sheets alphabetically by class name
        concrete_classes.sort(key=lambda cls: cls.name)
        
        wb = Workbook()
        wb.remove(wb.active)
        
        for cls in concrete_classes:
            print(f"Verwerken: {cls.name}")
            self._create_worksheet_for_class(wb, cls)
        
        wb.save(template_file_path)
        print()
        print("=" * 40)
        print("KLAAR")
        print(template_file_path)
        print("=" * 40)
    
    async def generate_infotemplate_from_subset_async(
            self,
            subset_path: Path,
            template_file_path: Path,
            model_directory: Path | None = None):
        """Async version of generate_infotemplate_from_subset."""
        self.collector = await self._load_collector_async(subset_path)
        self.classes = self.collector.classes
        self._build_relation_dict()
        
        concrete_classes = self._filter_concrete_classes()
        # Feature 12: Sort sheets alphabetically by class name
        concrete_classes.sort(key=lambda cls: cls.name)
        
        wb = Workbook()
        wb.remove(wb.active)
        
        for cls in concrete_classes:
            print(f"Verwerken: {cls.name}")
            self._create_worksheet_for_class(wb, cls)
            await asyncio.sleep(0)
        
        wb.save(template_file_path)
        print()
        print("=" * 40)
        print("KLAAR")
        print(template_file_path)
        print("=" * 40)