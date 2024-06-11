#!/usr/bin/env python3

import argparse
import pandas as pd
import re
from enum import Enum



parser = argparse.ArgumentParser()
parser.add_argument(
    "-i",
    "--input",
    help="Input Excel File",
    required=True,
)
parser.add_argument(
"-o",
    "--output",
    help="Output Markdown File",
    required=True,
)
parser.add_argument(
    "-V",
    "--version",
    help="Tisax Version",
    required=True,
)

args = parser.parse_args()

skiprows = None
sheet_name = None
nrows = None
column_indices = {}


match args.version:
    case "6_DE":
        skiprows = 1  # first row is a title
        sheet_name = 4  # 5th sheet
        nrows = None

        column_indices = {
            "controlnum": 2,  # C
            "controlquestion": 7,
            "goal": 8,
            "requirement_must": 9,
            "requirement_should": 10,
            "requirement_high": 11,
            "requirement_very_high": 12,
        }  # For ISA6_DE_6
    case "5_1_DE":
        skiprows = 1  # first row is a title
        sheet_name = 4  # 5th sheet
        nrows = 59

        column_indices = {
            "controlnum": 3,  # D
            "controlquestion": 8,
            "goal": 9,
            "requirement_must": 10,
            "requirement_should": 11,
            "requirement_high": 12,
            "requirement_very_high": 13,
        }  # For ISA6_DE_6
    case _:
        raise Exception("Sorry, Excel Columns not yet defined in Script. Only version 6_DE and 5_1_DE implemented.")

df = pd.read_excel(args.input, skiprows=skiprows, sheet_name=sheet_name, dtype=str, nrows=nrows)

# remove line breaks from column names
df.columns = df.columns.str.replace("\n", "")


def dataframe_to_markdown(df):
    markdown_lines = []
    for _, row in df.iterrows():
        levels = row[df.columns[column_indices["controlnum"]]].count(".") + 1
        header = "#" * levels
        markdown_lines.append(
            f"{header} {row[column_indices["controlnum"]]} {row[column_indices["controlquestion"]]}"
        )
        if levels > 2:
            description = f"""
**{df.columns[column_indices["goal"]]}**

{row[column_indices["goal"]]}

**{df.columns[column_indices["requirement_must"]]}**

{row[column_indices["requirement_must"]]}

**{df.columns[column_indices["requirement_should"]]}**

{row[column_indices["requirement_should"]]}

**{df.columns[column_indices["requirement_high"]]}**

{row[column_indices["requirement_high"]]}

**{df.columns[column_indices["requirement_very_high"]]}**

{row[column_indices["requirement_very_high"]]}
"""
            markdown_lines.append(description)

    return "\n".join(markdown_lines)

output = dataframe_to_markdown(df)

# replace all kinds of hyphens with ascii symbols
output = re.sub(r"[‐᠆﹣－⁃−–]+", "-", output)
# replace non breaking space with whitespace
output = re.sub(r"[ ]+", " ", output)
# make items starting with '-' subitems (some are not indented)
output = output.replace("\n-", "\n  -")

with open(args.output, "w") as f:
    f.write(output)
