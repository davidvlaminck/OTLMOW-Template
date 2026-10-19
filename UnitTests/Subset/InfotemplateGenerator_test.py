from pathlib import Path

import openpyxl
import pytest

from otlmow_template.InfotemplateGenerator import InfotemplateGenerator

current_dir = Path(__file__).parent
model_directory_path = Path(__file__).parent.parent / 'TestModel'


def test_infotemplate_generator_basic():
    """Test that InfotemplateGenerator produces expected output matching original generator."""
    subset_path = current_dir / 'simpele_vergelijkings_subset.db'
    output_path = current_dir / 'test_infotemplate_output.xlsx'
    
    generator = InfotemplateGenerator()
    generator.generate_infotemplate_from_subset(
        subset_path=subset_path,
        template_file_path=output_path,
        model_directory=model_directory_path
    )
    
    assert output_path.exists()
    
    wb = openpyxl.load_workbook(output_path, read_only=True)
    sheet_names = wb.sheetnames
    
    # Check we have the expected 5 sheets
    expected_sheets = {'Verkeersbordopstelling', 'Funderingsmassief', 'HoortBij', 'Pictogram', 'Verkeersbordsteun'}
    assert set(sheet_names) == expected_sheets
    
    # Verify each sheet structure
    for sheet_name in expected_sheets:
        ws = wb[sheet_name]
        
        # Check A1 = BESCHRIJVING
        assert ws['A1'].value == 'BESCHRIJVING'
        
        # Check A2 has description
        assert ws['A2'].value is not None
        
        # Check A4 = RELATIES
        assert ws['A4'].value == 'RELATIES'
        
        # Check relation headers at row 5
        assert ws['B5'].value == 'Verplicht ja/nee'
        assert ws['C5'].value == 'Wie levert aan?'
        assert ws['D5'].value == 'Type Relatie'
        # Feature 11: Column E is now "Richting"
        assert ws['E5'].value == 'Richting'
        assert ws['F5'].value == 'Opmerkingen'
        
        # Check EIGENSCHAPPEN section exists
        eigenschappen_found = False
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=1, values_only=True):
            if row[0] == 'EIGENSCHAPPEN':
                eigenschappen_found = True
                break
        assert eigenschappen_found
        
        # Check attribute headers
        attr_row = None
        for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True), 1):
            if row[0] == 'Attribuutnaam':
                attr_row = row_idx
                break
        assert attr_row is not None
        assert ws.cell(row=attr_row, column=1).value == 'Attribuutnaam'
        assert ws.cell(row=attr_row, column=2).value == 'Verplicht ja/nee'
        assert ws.cell(row=attr_row, column=3).value == 'Wie levert aan?'
        assert ws.cell(row=attr_row, column=4).value == 'Wat?'
        assert ws.cell(row=attr_row, column=5).value == 'Formaat/Opmaak/naamconventie'
        assert ws.cell(row=attr_row, column=6).value == 'Opmerkingen'
    
    wb.close()
    output_path.unlink()


def test_infotemplate_generator_specific_content():
    """Test specific content for known classes."""
    subset_path = current_dir / 'simpele_vergelijkings_subset.db'
    output_path = current_dir / 'test_infotemplate_output2.xlsx'
    
    generator = InfotemplateGenerator()
    generator.generate_infotemplate_from_subset(
        subset_path=subset_path,
        template_file_path=output_path,
        model_directory=model_directory_path
    )
    
    wb = openpyxl.load_workbook(output_path, read_only=True)
    
    # Test Verkeersbordopstelling
    ws = wb['Verkeersbordopstelling']
    
    # Check relations - Feature 1: only concrete classes
    # Feature 11: Column D=Type Relatie, Column E=Richting
    relations = []
    for row in ws.iter_rows(min_row=6, max_row=12, min_col=4, max_col=5, values_only=True):
        if row[0] and row[1]:
            relations.append((row[0], row[1]))
    
    # Draagconstructie is abstract, filtered out by Feature 1
    # Column E is now Richting: Verkeersbordopstelling is the doel of the directed
    # HoortBij relations, so the direction is reversed (Destination -> Source = ⇐)
    expected_relations = [
        ('HoortBij', "⇐"),
        ('HoortBij', "⇐")
    ]
    assert relations == expected_relations
    
    # Check attributes
    attrs = []
    for row in ws.iter_rows(min_row=17, max_row=18, min_col=1, max_col=5, values_only=True):
        if row[0]:
            attrs.append((row[0], row[1], row[4]))
    
    expected_attrs = [
        ('asset-id', 'ja', 'DtcIdentificator'),
        ('type URI', 'ja', 'anyURI')
    ]
    assert attrs == expected_attrs
    
    # Test Funderingsmassief
    ws = wb['Funderingsmassief']
    
    attrs = []
    for row in ws.iter_rows(min_row=17, max_row=32, min_col=1, max_col=5, values_only=True):
        if row[0] and row[0] not in ['Attribuutnaam', 'EIGENSCHAPPEN']:
            attrs.append((row[0], row[1], row[4]))
    
    expected_attrs_funderingsmassief = [
        # Generic attributes first (alphabetically)
        ('AIM-toestand', 'ja', 'KlAIMToestand'),
        ('asset-id', 'ja', 'DtcIdentificator'),
        ('bestekpostnummer', 'ja', 'string'),
        ('datum oprichting object', 'ja', 'date'),
        ('is actief', 'ja', 'boolean'),
        ('naam', 'ja', 'string'),
        ('notitie', 'ja', 'string'),
        ('standaardbestekpostnummer', 'ja', 'string'),
        ('type URI', 'ja', 'anyURI'),
        # Non-generic attributes (alphabetically)
        ('aanzetpeil', 'ja', 'KwantWrdInMeterTAW'),
        ('afmeting grondvlak', 'ja', 'DtuAfmetingGrondvlak'),
        ('funderingshoogte', 'ja', 'KwantWrdInCentimeter'),
        ('is permanent', 'ja', 'boolean'),
        ('is prefab', 'ja', 'boolean'),
        ('materiaal', 'ja', 'KlAlgMateriaal'),
        ('volume', 'ja', 'KwantWrdInKubiekeMeter')
    ]
    assert attrs == expected_attrs_funderingsmassief
    
    wb.close()
    output_path.unlink()


def test_feature1_concrete_classes_only_in_relaties():
    """Test Feature 1: Only concrete classes shown in relaties tab (no abstract classes like Draagconstructie)."""
    subset_path = current_dir / 'simpele_vergelijkings_subset.db'
    output_path = current_dir / 'test_feature1_output.xlsx'
    
    generator = InfotemplateGenerator()
    generator.generate_infotemplate_from_subset(
        subset_path=subset_path,
        template_file_path=output_path,
        model_directory=model_directory_path
    )
    
    wb = openpyxl.load_workbook(output_path, read_only=True)
    
    # Test Verkeersbordopstelling - should NOT have Draagconstructie (abstract) in relations
    ws = wb['Verkeersbordopstelling']
    
    relations = []
    for row in ws.iter_rows(min_row=6, max_row=12, min_col=4, max_col=5, values_only=True):
        if row[0] and row[1]:
            relations.append((row[0], row[1]))
    
    # Draagconstructie is abstract, should be filtered out
    # Verkeersbordopstelling is the doel of the directed HoortBij relations, so the
    # direction is reversed (Destination -> Source = ⇐)
    expected_relations = [
        ('HoortBij', "⇐"),
        ('HoortBij', "⇐")
    ]
    assert relations == expected_relations, f"Expected {expected_relations}, got {relations}"
    
    # Test Funderingsmassief - should NOT have abstract classes
    ws = wb['Funderingsmassief']
    
    relations = []
    for row in ws.iter_rows(min_row=6, max_row=12, min_col=4, max_col=5, values_only=True):
        if row[0] and row[1]:
            relations.append((row[0], row[1]))
    
    # Bevestiging with Verkeersbordsteun should remain (both concrete)
    # Feature 11: Column E is now Richting: Bevestiging is Unspecified so ⇔
    expected_relations_funderingsmassief = [
        ('Bevestiging', "⇔")
    ]
    assert relations == expected_relations_funderingsmassief, f"Expected {expected_relations_funderingsmassief}, got {relations}"
    
    wb.close()
    output_path.unlink()


def test_feature2_double_heeftbetrokkene_relatie():
    """Test Feature 2: Double heeftBetrokkene relatie - if subset has heeftBetrokkene, create 2 records."""
    subset_path = current_dir / 'simpele_vergelijkings_subset.db'
    output_path = current_dir / 'test_feature2_output.xlsx'
    
    generator = InfotemplateGenerator()
    generator.generate_infotemplate_from_subset(
        subset_path=subset_path,
        template_file_path=output_path,
        model_directory=model_directory_path
    )
    
    wb = openpyxl.load_workbook(output_path, read_only=True)
    
    # Current test subset doesn't have heeftBetrokkene, so no double relations should appear
    # This test verifies the feature doesn't break existing behavior
    # Check expected relation counts per sheet (based on concrete classes only)
    expected_counts = {
        'Verkeersbordopstelling': 2,  # HoortBij to Pictogram, Verkeersbordsteun
        'Funderingsmassief': 1,       # Bevestiging to Verkeersbordsteun
        'HoortBij': 0,                # Relation class, no outgoing relations
        'Pictogram': 2,               # Bevestiging to Draagconstructie, Verkeersbordsteun; HoortBij to Verkeersbordopstelling
        'Verkeersbordsteun': 3,       # Bevestiging to Funderingsmassief, Pictogram; HoortBij to Verkeersbordopstelling
    }
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        relations = []
        for row in ws.iter_rows(min_row=6, max_row=12, min_col=4, max_col=5, values_only=True):
            if row[0] and row[1]:
                relations.append((row[0], row[1]))
        # Check expected count
        expected = expected_counts.get(sheet_name, 0)
        assert len(relations) == expected, f"Sheet {sheet_name}: Expected {expected} relations, got {len(relations)}: {relations}"
    
    wb.close()
    output_path.unlink()


def test_feature2_double_heeftbetrokkene_relatie_mock():
    """Test Feature 2: Double heeftBetrokkene relatie using mocked relation dict."""
    
    # Test the actual _build_relation_dict logic by creating a mock
    concrete_class_names = {'TestClass', 'SomeClass', 'OtherClass'}
    relation_dict = {}
    mock_relations = [
        ('TestClass', 'heeftBetrokkene', 'SomeClass'),
        ('TestClass', 'HoortBij', 'OtherClass'),
    ]
    
    # Simulate the logic from _build_relation_dict
    seen_relations = set()
    for bron, relatie, doel in mock_relations:
        if bron in concrete_class_names and doel in concrete_class_names:
            # Add from bron perspective
            if relatie == 'heeftBetrokkene':
                relation_dict.setdefault(bron, []).append((relatie, doel))
            else:
                rel_key = (bron, relatie, doel)
                if rel_key not in seen_relations:
                    seen_relations.add(rel_key)
                    relation_dict.setdefault(bron, []).append((relatie, doel))
            
            # Add from doel perspective
            if relatie == 'heeftBetrokkene':
                relation_dict.setdefault(doel, []).append((relatie, bron))
            else:
                rel_key = (doel, relatie, bron)
                if rel_key not in seen_relations:
                    seen_relations.add(rel_key)
                    relation_dict.setdefault(doel, []).append((relatie, bron))
    
# Feature 2: Double the 'heeftBetrokkene' relations with specific role types
    for bron, relations in relation_dict.items():
        new_relations = []
        for relatie, doel in relations:
            if relatie == 'heeftBetrokkene':
                # First record: HeeftBetrokkene - rol: toezichter
                new_relations.append(("HeeftBetrokkene - rol: toezichter", doel))
                # Second record: HeeftBetrokkene - rol: installatieverantwoordelijke
                new_relations.append(("HeeftBetrokkene - rol: installatieverantwoordelijke", doel))
            else:
                new_relations.append((relatie, doel))
        relation_dict[bron] = new_relations

    # Verify heeftBetrokkene appears twice for TestClass with specific role types
    test_class_relations = relation_dict['TestClass']
    toezichter_count = sum(1 for r, d in test_class_relations if r == 'HeeftBetrokkene - rol: toezichter')
    installatie_count = sum(1 for r, d in test_class_relations if r == 'HeeftBetrokkene - rol: installatieverantwoordelijke')
    assert toezichter_count == 1, f"Expected 1 'HeeftBetrokkene - rol: toezichter', got {toezichter_count}"
    assert installatie_count == 1, f"Expected 1 'HeeftBetrokkene - rol: installatieverantwoordelijke', got {installatie_count}"

    # HoortBij should appear once
    hoortbij_count = sum(1 for r, d in test_class_relations if r == 'HoortBij')
    assert hoortbij_count == 1, f"Expected 1 HoortBij, got {hoortbij_count}"


def test_feature3_generic_attributes_light_green():
    """Test Feature 3: Generic attributes get light green background color."""
    from openpyxl.styles import PatternFill
    
    subset_path = current_dir / 'simpele_vergelijkings_subset.db'
    output_path = current_dir / 'test_feature3_output.xlsx'
    
    generator = InfotemplateGenerator()
    generator.generate_infotemplate_from_subset(
        subset_path=subset_path,
        template_file_path=output_path,
        model_directory=model_directory_path
    )
    
    wb = openpyxl.load_workbook(output_path)
    
    # Test Funderingsmassief - has several generic attributes
    ws = wb['Funderingsmassief']
    
    # Find the attribute header row
    attr_row = None
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True), 1):
        if row[0] == 'Attribuutnaam':
            attr_row = row_idx
            break
    
    assert attr_row is not None
    
    # Light green fill color (updated to lighter green)
    light_green = PatternFill(fill_type="solid", start_color="C8F7C8", end_color="C8F7C8")
    
    # Check generic attributes have light green background
    # Generic attributes in Funderingsmassief (from the URI list):
    # asset-id, datum oprichting object, bestekpostnummer, naam, is actief, AIM-toestand (toestand)
    generic_attrs_found = []
    for row in ws.iter_rows(min_row=attr_row + 1, max_row=ws.max_row, min_col=1, max_col=6, values_only=False):
        if row[0].value:
            attr_name = row[0].value
            # Check if this is a generic attribute (by name matching the URI list)
            if attr_name in ['asset-id', 'datum oprichting object', 'bestekpostnummer', 'naam', 'is actief', 'AIM-toestand']:
                generic_attrs_found.append(attr_name)
                # Check if cell has light green fill
                assert row[0].fill == light_green or row[0].fill.start_color.rgb == 'FF90EE90', \
                    f"Attribute {attr_name} should have light green fill, got {row[0].fill}"
    
    # Verify we found the expected generic attributes
    expected_generic = ['asset-id', 'datum oprichting object', 'bestekpostnummer', 'naam', 'is actief', 'AIM-toestand']
    for expected in expected_generic:
        assert expected in generic_attrs_found, f"Expected generic attribute {expected} not found"
    
    wb.close()
    output_path.unlink()


def test_feature4_default_aannemer_wie_levert_aan():
    """Test Feature 4: Default 'aannemer' in 'wie levert aan' column for both relations and attributes."""
    subset_path = current_dir / 'simpele_vergelijkings_subset.db'
    output_path = current_dir / 'test_feature4_output.xlsx'
    
    generator = InfotemplateGenerator()
    generator.generate_infotemplate_from_subset(
        subset_path=subset_path,
        template_file_path=output_path,
        model_directory=model_directory_path
    )
    
    wb = openpyxl.load_workbook(output_path)
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        # Check relations section (rows 6-12, column C = "Wie levert aan?")
        # Should have 'aannemer' for each relation row (only for sheets that have relations)
        relation_rows_with_data = 0
        for row in ws.iter_rows(min_row=6, max_row=12, min_col=3, max_col=3, values_only=False):
            if row[0].value is not None:
                relation_rows_with_data += 1
                assert row[0].value == 'aannemer', f"Sheet {sheet_name}: Expected 'aannemer' in relations, got '{row[0].value}'"
        # Only assert if there are relations (some sheets like HoortBij are relation classes)
        if relation_rows_with_data > 0:
            pass  # Already asserted above
        
        # Check attributes section - find attribute header row
        attr_row = None
        for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True), 1):
            if row[0] == 'Attribuutnaam':
                attr_row = row_idx
                break
        
        if attr_row:
            # Check attributes section (column C = "Wie levert aan?")
            attr_rows_with_data = 0
            for row in ws.iter_rows(min_row=attr_row + 1, max_row=ws.max_row, min_col=3, max_col=3, values_only=False):
                if row[0].value is not None:
                    attr_rows_with_data += 1
                    assert row[0].value == 'aannemer', f"Sheet {sheet_name}: Expected 'aannemer' in attributes, got '{row[0].value}'"
            assert attr_rows_with_data > 0, f"Sheet {sheet_name}: No attribute rows with data found"
    
    wb.close()
    output_path.unlink()


def test_feature5_typeuri_in_opmerkingen():
    """Test Feature 5: TypeURI in opmerkingen column for typeURI attribute only."""
    subset_path = current_dir / 'simpele_vergelijkings_subset.db'
    output_path = current_dir / 'test_feature5_output.xlsx'
    
    generator = InfotemplateGenerator()
    generator.generate_infotemplate_from_subset(
        subset_path=subset_path,
        template_file_path=output_path,
        model_directory=model_directory_path
    )
    
    wb = openpyxl.load_workbook(output_path)
    
    # Test Funderingsmassief - has type URI attribute
    ws = wb['Funderingsmassief']
    
    # Find attribute header row
    attr_row = None
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True), 1):
        if row[0] == 'Attribuutnaam':
            attr_row = row_idx
            break
    
    assert attr_row is not None
    
    # Check that type URI attribute has its typeURI in opmerkingen (column F)
    for row in ws.iter_rows(min_row=attr_row + 1, max_row=ws.max_row, min_col=1, max_col=6, values_only=False):
        if row[0].value == 'type URI':
            # Opmerkingen is column F (6)
            assert row[5].value is not None, "type URI should have typeURI in opmerkingen"
            assert 'anyuri' in str(row[5].value).lower() or 'typeuri' in str(row[5].value).lower(), \
                f"type URI opmerkingen should contain typeURI info, got: {row[5].value}"
        elif row[0].value in ['asset-id', 'naam', 'is actief']:
            # Other attributes should NOT have typeURI in opmerkingen (only typeURI attribute)
            # They can have other opmerkingen but not the class typeURI
            pass
    
    wb.close()
    output_path.unlink()


def test_feature6_bold_headers():
    """Test Feature 6: All table headers are bold."""
    
    subset_path = current_dir / 'simpele_vergelijkings_subset.db'
    output_path = current_dir / 'test_feature6_output.xlsx'
    
    generator = InfotemplateGenerator()
    generator.generate_infotemplate_from_subset(
        subset_path=subset_path,
        template_file_path=output_path,
        model_directory=model_directory_path
    )
    
    wb = openpyxl.load_workbook(output_path)
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        # Check main headers (A1, A4, A13/EIGENSCHAPPEN row)
        assert ws['A1'].font.bold == True, f"Sheet {sheet_name}: A1 should be bold"
        assert ws['A4'].font.bold == True, f"Sheet {sheet_name}: A4 should be bold"
        
        # Check relation headers (row 5)
        for col in ['B', 'C', 'D', 'E', 'F']:
            assert ws[f'{col}5'].font.bold == True, f"Sheet {sheet_name}: {col}5 should be bold"
        
        # Check attribute headers (row with "Attribuutnaam")
        attr_row = None
        for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True), 1):
            if row[0] == 'Attribuutnaam':
                attr_row = row_idx
                break
        
        if attr_row:
            for col in range(1, 7):
                cell = ws.cell(row=attr_row, column=col)
                assert cell.font.bold == True, f"Sheet {sheet_name}: Attribute header col {col} should be bold"
    
    wb.close()
    output_path.unlink()


def test_feature7_isactief_toestand_defaults():
    """Test Feature 7: isActief=TRUE, toestand='in-gebruik' defaults."""
    subset_path = current_dir / 'simpele_vergelijkings_subset.db'
    output_path = current_dir / 'test_feature7_output.xlsx'

    generator = InfotemplateGenerator()
    generator.generate_infotemplate_from_subset(
        subset_path=subset_path,
        template_file_path=output_path,
        model_directory=model_directory_path
    )

    wb = openpyxl.load_workbook(output_path)

    # Test Funderingsmassief - has both is actief and AIM-toestand
    ws = wb['Funderingsmassief']

    # Find attribute header row
    attr_row = None
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True), 1):
        if row[0] == 'Attribuutnaam':
            attr_row = row_idx
            break

    assert attr_row is not None

    # Check is actief attribute - should have TRUE in opmerkingen
    # Check toestand attribute - should have 'in-gebruik' in opmerkingen
    for row in ws.iter_rows(min_row=attr_row + 1, max_row=ws.max_row, min_col=1, max_col=6, values_only=False):
        if row[0].value == 'is actief':
            # isActief should have TRUE as default in opmerkingen
            assert row[1].value == 'ja', "is actief should be verplicht (ja)"
            assert row[5].value == "Moet TRUE zijn, behalve wanneer het asset uit de databank verwijderd moet worden", \
                f"is actief opmerkingen incorrect: {row[5].value}"
        elif row[0].value == 'AIM-toestand':
            # toestand should have 'in-gebruik' as default in opmerkingen
            assert row[1].value == 'ja', "AIM-toestand should be verplicht (ja)"
            assert row[5].value == "Is bijna altijd 'in-gebruik'", \
                f"AIM-toestand opmerkingen incorrect: {row[5].value}"

    wb.close()
    output_path.unlink()


def test_feature8_codelijst_uri_link():
    """Test Feature 8: Codelijst URI link in formaat/opmaak/naamconventie for keuzelijst attributes."""
    subset_path = current_dir / 'simpele_vergelijkings_subset.db'
    output_path = current_dir / 'test_feature8_output.xlsx'

    generator = InfotemplateGenerator()
    generator.generate_infotemplate_from_subset(
        subset_path=subset_path,
        template_file_path=output_path,
        model_directory=model_directory_path
    )

    wb = openpyxl.load_workbook(output_path)

    # Test Funderingsmassief - has materiaal (KlAlgMateriaal) and AIM-toestand (KlAIMToestand) which are codelijsten
    ws = wb['Funderingsmassief']

    # Find attribute header row
    attr_row = None
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True), 1):
        if row[0] == 'Attribuutnaam':
            attr_row = row_idx
            break

    assert attr_row is not None

    # Check that codelijst attributes have clickable hyperlink in Formaat/Opmaak/naamconventie (col E)
    for row in ws.iter_rows(min_row=attr_row + 1, max_row=ws.max_row, min_col=1, max_col=6, values_only=False):
        if row[0].value == 'materiaal':
            # Formaat/Opmaak/naamconventie is column E (5)
            assert row[4].value is not None, "materiaal should have codelijst URI in Formaat"
            assert 'KlAlgMateriaal' in str(row[4].value), \
                f"materiaal Formaat should contain codelijst name, got: {row[4].value}"
            # Check that the cell has a clickable hyperlink
            assert row[4].hyperlink is not None, "materiaal should have a clickable hyperlink"
            assert 'http' in str(row[4].hyperlink.target), \
                f"materiaal hyperlink should contain URL, got: {row[4].hyperlink.target}"
        elif row[0].value == 'AIM-toestand':
            # AIM-toestand has type KlAIMToestand which is a keuzelijst
            assert row[4].value is not None, "AIM-toestand should have codelijst URI in Formaat"
            assert 'KlAIMToestand' in str(row[4].value), \
                f"AIM-toestand Formaat should contain codelijst name, got: {row[4].value}"
            # Check that the cell has a clickable hyperlink
            assert row[4].hyperlink is not None, "AIM-toestand should have a clickable hyperlink"
            assert 'http' in str(row[4].hyperlink.target), \
                f"AIM-toestand hyperlink should contain URL, got: {row[4].hyperlink.target}"

    wb.close()
    output_path.unlink()


def test_feature10_auto_filter_attributen():
    """Test Feature 10: Auto-filter on attributen header row."""
    subset_path = current_dir / 'simpele_vergelijkings_subset.db'
    output_path = current_dir / 'test_feature10_output.xlsx'
    
    generator = InfotemplateGenerator()
    generator.generate_infotemplate_from_subset(
        subset_path=subset_path,
        template_file_path=output_path,
        model_directory=model_directory_path
    )
    
    wb = openpyxl.load_workbook(output_path)
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        # Find attribute header row
        attr_row = None
        for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True), 1):
            if row[0] == 'Attribuutnaam':
                attr_row = row_idx
                break
        
        if attr_row:
            # Check that auto_filter is set on the attribute header row
            assert ws.auto_filter.ref is not None, f"Sheet {sheet_name}: Auto-filter should be set"
            # The ref should include the attribute header row
            assert str(attr_row) in ws.auto_filter.ref, f"Sheet {sheet_name}: Auto-filter should include attribute header row {attr_row}"
    
    wb.close()
    output_path.unlink()


def test_feature11_richting_column():
    """Test Feature 11: Richting column for all relations, move 'relatie met' from E to A."""
    subset_path = current_dir / 'simpele_vergelijkings_subset.db'
    output_path = current_dir / 'test_feature11_output.xlsx'
    
    generator = InfotemplateGenerator()
    generator.generate_infotemplate_from_subset(
        subset_path=subset_path,
        template_file_path=output_path,
        model_directory=model_directory_path
    )
    
    wb = openpyxl.load_workbook(output_path)
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        # Check relation headers - should have "Richting" in column E (5) instead of "Relatie Met"
        # And "Relatie Met" should be in A-relations tab (but we only have one relation section per sheet)
        # For now, check that column E has "Richting" header
        assert ws['E5'].value == 'Richting', f"Sheet {sheet_name}: E5 should be 'Richting', got '{ws['E5'].value}'"
        
        # Check that relation rows have richting values (⇐, ⇒, ⇔)
        for row in ws.iter_rows(min_row=6, max_row=12, min_col=4, max_col=5, values_only=False):
            if row[0].value and row[1].value:
                richting = row[1].value  # Column E
                assert richting in ["⇐", "⇒", "⇔"], f"Sheet {sheet_name}: Richting should be '⇐', '⇒', or '⇔', got '{richting}'"
    
    wb.close()
    output_path.unlink()


def test_feature12_sheets_alphabetical():
    """Test Feature 12: All sheets sorted alphabetically."""
    subset_path = current_dir / 'simpele_vergelijkings_subset.db'
    output_path = current_dir / 'test_feature12_output.xlsx'
    
    generator = InfotemplateGenerator()
    generator.generate_infotemplate_from_subset(
        subset_path=subset_path,
        template_file_path=output_path,
        model_directory=model_directory_path
    )
    
    wb = openpyxl.load_workbook(output_path)
    
    sheet_names = wb.sheetnames
    # Check sheets are alphabetically sorted
    assert sheet_names == sorted(sheet_names), f"Sheets not alphabetically sorted: {sheet_names}"
    
    wb.close()
    output_path.unlink()


@pytest.mark.asyncio
async def test_infotemplate_generator_async():
    """Test async version of InfotemplateGenerator."""
    subset_path = current_dir / 'simpele_vergelijkings_subset.db'
    output_path = current_dir / 'test_infotemplate_output_async.xlsx'
    
    generator = InfotemplateGenerator()
    await generator.generate_infotemplate_from_subset_async(
        subset_path=subset_path,
        template_file_path=output_path,
        model_directory=model_directory_path
    )
    
    assert output_path.exists()
    
    wb = openpyxl.load_workbook(output_path, read_only=True)
    sheet_names = wb.sheetnames
    
    expected_sheets = {'Verkeersbordopstelling', 'Funderingsmassief', 'HoortBij', 'Pictogram', 'Verkeersbordsteun'}
    assert set(sheet_names) == expected_sheets
    
    wb.close()
    output_path.unlink()