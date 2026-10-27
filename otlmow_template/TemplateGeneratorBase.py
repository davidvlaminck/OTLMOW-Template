import asyncio
import sqlite3
from pathlib import Path
from typing import Optional, List

from otlmow_modelbuilder.OSLOCollector import OSLOCollector
from otlmow_modelbuilder.SQLDataClasses.OSLOClass import OSLOClass


class TemplateGeneratorBase:
    """Base class for template generators providing common functionality."""
    
    def __init__(self):
        self.collector: Optional[OSLOCollector] = None
        self.classes: List[OSLOClass] = []
    
    def _load_collector_from_subset_path(self, subset_path: Path) -> OSLOCollector:
        """Load and initialize OSLOCollector from subset file."""
        collector = OSLOCollector(subset_path)
        collector.collect_all(include_abstract=True)
        return collector
    
    @classmethod
    def _filter_classes_by_uris(cls, classes: List[OSLOClass], class_uris_filter: List[str] = None) -> List[OSLOClass]:
        """Filter classes by URIs. If None, return all classes."""
        if class_uris_filter is None:
            return classes
        return [x for x in classes if x.objectUri in class_uris_filter]
    
    def _is_relation_class(self, class_uri: str) -> bool:
        """Determine if a class is a relation class by walking up the inheritance chain."""
        if not self.collector or not getattr(self.collector, 'path', None):
            return False
        try:
            conn = sqlite3.connect(self.collector.path)
            cur = conn.cursor()
            visited = set()
            current = class_uri
            depth = 0
            while current and current not in visited and depth < 3:
                visited.add(current)
                cur.execute("SELECT base_name, base_uri FROM InternalBaseClass WHERE class_uri = ?", (current,))
                rows = cur.fetchall()
                if not rows:
                    break
                moved = False
                for base_name, base_uri in rows:
                    base_name_l = (base_name or '').lower()
                    if base_name_l and ('relatieobject' in base_name_l or base_name in ('DirectioneleRelatie', 'NietDirectioneleRelatie')):
                        conn.close()
                        return True
                    if base_uri and base_uri not in visited:
                        current = base_uri
                        depth += 1
                        moved = True
                        break
                if not moved:
                    break
            conn.close()
            return False
        except Exception:
            return False
    
    def _filter_concrete_classes(self) -> List[OSLOClass]:
        """Filter classes to only include concrete (non-abstract) classes."""
        concrete_classes = []
        for cls in self.classes:
            if cls.abstract == 0:
                if self._is_relation_class(cls.objectUri):
                    continue
                concrete_classes.append(cls)
        return concrete_classes
    
    def _get_all_concrete_classes(self, class_uris_filter: List[str] = None) -> List[OSLOClass]:
        """Get all concrete classes, optionally filtered by URIs."""
        filtered = self._filter_classes_by_uris(self.classes, class_uris_filter)
        concrete = []
        for cls in filtered:
            if cls.abstract == 0:
                if self._is_relation_class(cls.objectUri):
                    continue
                concrete.append(cls)
        return concrete
    
    async def _load_collector_async(self, subset_path: Path) -> OSLOCollector:
        """Async version of loading collector."""
        await asyncio.sleep(0)
        collector = self._load_collector_from_subset_path(subset_path)
        await asyncio.sleep(0)
        return collector