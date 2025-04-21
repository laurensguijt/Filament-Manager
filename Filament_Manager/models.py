from dataclasses import dataclass
from datetime import datetime
from typing import Optional


@dataclass
class FilamentData:
    code: str
    material: str
    variant: str
    supplier: str
    date_opened: datetime
    weight: float
    hex_color: str
    empty_spool_weight: float
    description: Optional[str] = ""

    @classmethod
    def from_row(cls, row):
        """Create a FilamentData instance from an Excel row"""
        return cls(
            code=str(row[0]),
            material=str(row[1]),
            variant=str(row[2]),
            supplier=str(row[3]),
            date_opened=row[4],
            weight=float(row[5]) if row[5] is not None else 0,
            hex_color=str(row[6]) if len(row) > 6 and isinstance(row[6], str) and row[6].startswith('#') else "#000000",
            empty_spool_weight=float(row[7]) if len(row) > 7 and row[7] is not None else 0,
            description=str(row[8]) if len(row) > 8 and row[8] is not None else ""
        )

    def to_row(self):
        """Convert the FilamentData instance to a row for Excel"""
        return [
            self.code,
            self.material,
            self.variant,
            self.supplier,
            self.date_opened,
            self.weight,
            self.hex_color,
            self.empty_spool_weight,
            self.description
        ]


@dataclass
class PrintLogEntry:
    timestamp: str
    print_name: str
    filament_code: str
    material: str
    variant: str
    used_weight: float
    remaining_weight: float

    @classmethod
    def from_row(cls, row):
        """Create a PrintLogEntry instance from an Excel row"""
        return cls(
            timestamp=str(row[0]),
            print_name=str(row[1]),
            filament_code=str(row[2]),
            material=str(row[3]),
            variant=str(row[4]),
            used_weight=float(row[5]) if row[5] is not None else 0,
            remaining_weight=float(row[6]) if row[6] is not None else 0
        )

    def to_row(self):
        """Convert the PrintLogEntry instance to a row for Excel"""
        return [
            self.timestamp,
            self.print_name,
            self.filament_code,
            self.material,
            self.variant,
            self.used_weight,
            self.remaining_weight
        ] 