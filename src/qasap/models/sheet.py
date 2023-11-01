from dataclasses import dataclass, field, InitVar, FrozenInstanceError
from typing import Any, List
from enum import Enum

from pandas import DataFrame
import streamlit as st


@dataclass(frozen=True)
class Sheet:
    original_sheet: DataFrame


@dataclass(frozen=True)
class Introduction(Sheet):
    name: str = "Introduction"

    def __post_init__(self):
        col3_data = self.original_sheet.iloc[:, 2].dropna().tolist()
        object.__setattr__(self, "lines", col3_data)

    def render(self):
        with st.expander("Introduction"):
            for line in self.lines:
                st.write(line)


class FieldRequirement(Enum):
    Mandatory = "Mandatory"
    Optional = "Optional"
    Blank = "Blank"


class FieldType(Enum):
    Text = "Text"
    Number = "Number"
    Date = "Date"


char_to_type = {
    "C": FieldType.Text,
    "P": FieldType.Number,
    "D": FieldType.Date,
}


@dataclass(frozen=True)
class FieldList(Sheet):
    name: str = "Field List"


@dataclass(frozen=False)
class Field:
    name: str
    technical_name: str
    type: str
    length: int
    decimals: int
    description: str
    requirement: FieldRequirement = FieldRequirement.Blank
    values: List[str] = field(default_factory=list)


class MasterDetails(Sheet):
    def __init__(self, original_sheet, fields_requirements: List[FieldRequirement] = []):
        super().__init__(original_sheet)
        self.name = "Master Details"
        self.version = self.original_sheet.iloc[0, 0]
        fields_technical_names = self.original_sheet.loc[3].tolist()
        fields_blobs = self.original_sheet.loc[4].tolist()
        fields_descriptions = self.original_sheet.loc[6].tolist()

        self.fields = []

        for idx, (tech_name, blob, desc) in enumerate(
            zip(fields_technical_names, fields_blobs, fields_descriptions)
        ):
            _, _, _, type_char, length, decimals = blob.split(";")
            type = char_to_type[type_char]
            name, description = desc.split("\n", 1)
            requirement = FieldRequirement.Mandatory if "*" in name else FieldRequirement.Blank
            name = name[:-1] if "*" in name else name

            # Get values from row 7 onward for the current column (field)
            values = self.original_sheet.iloc[7:, idx].dropna().tolist()

            self.fields.append(
                Field(
                    name,
                    tech_name,
                    type,
                    int(length),
                    int(decimals),
                    description,
                    requirement,
                    values,
                )
            )

        # TODO add key: bool to each field (riga 5)

    def render(self):
        st.write(f"## {self.name}")
        st.write(self.version)
        for f in self.fields:
            st.write(f)


EXTRA_SHEETS = [Introduction, FieldList]
