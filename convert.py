#!/usr/bin/env python3

import pdb
import argparse
import pandas as pd
import re


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
parser.add_argument(
    "-p",
    "--prototype",
    action="store_true",
    help="Export Controls for Prototype Protection",
    required=False,
)
parser.add_argument(
    "-d",
    "--data_protection",
    action="store_true",
    help="Export Controls for Data Protection",
    required=False,
)

args = parser.parse_args()


def c2int(char: str) -> int:
    return ord(char) - 64


excel_inidces = {
    "skiprows": 1,  # first row is a title
    "sheet_infosec": 4,  # 5th sheet
    "nrows_infosec": 59,
    "sheet_prototype": 5,
    "sheet_data_protection": 6,
    "nrows_data_protection": 5,
    "nrows_prototype": 30,
    "controlnum": 3,  # D
    "controlquestion": 8,
    "goal": ord("I") - 64,
    "requirement_must": 10,
    "requirement_should": 11,
    "requirement_high": 12,
    "requirement_very_high": 13,
}  # For ISA6_DE_6


match args.version:
    case "6_DE":

        excel_inidces = {
            "skiprows": 1,  # first row is a title
            "sheet_infosec": 4,  # 5th sheet
            "nrows_infosec": None,
            "sheet_prototype": 5,
            "nrows_prototype": None,
            "sheet_data_protection": 6,
            "nrows_data_protection": 21,
            "controlnum": 2,  # C
            "controlquestion": 7,
            "goal": 8,
            "requirement_must": 9,
            "requirement_should": 10,
            "requirement_high": 11,
            "requirement_very_high": 12,
        }  # For ISA6_DE_6
    case "5_1_DE":
        pass  # default dict
    case _:
        raise Exception(
            "Sorry, Excel Columns not yet defined in Script. Only version 6_DE and 5_1_DE implemented."
        )


def read_infosec(excel_inidces: dict) -> pd.DataFrame:
    df = pd.read_excel(
        args.input,
        skiprows=excel_inidces["skiprows"],
        sheet_name=excel_inidces["sheet_infosec"],
        nrows=excel_inidces["nrows_infosec"],
        dtype=str,
    )
    assert type(df) == pd.DataFrame
    df.columns = df.columns.str.replace("\n", "")
    return df


def read_prototype(excel_inidces: dict) -> pd.DataFrame:
    df_prototype = pd.read_excel(
        args.input,
        skiprows=excel_inidces["skiprows"],
        sheet_name=excel_inidces["sheet_prototype"],
        nrows=excel_inidces["nrows_prototype"],
        dtype=str,
    )
    assert type(df_prototype) == pd.DataFrame
    df_prototype.columns = df_prototype.columns.str.replace("\n", "")
    return df_prototype


def read_data_protection(excel_inidces: dict) -> pd.DataFrame:
    df_data_protection = pd.read_excel(
        args.input,
        skiprows=excel_inidces["skiprows"],
        sheet_name=excel_inidces["sheet_data_protection"],
        nrows=excel_inidces["nrows_data_protection"],
        dtype=str,
    )
    assert type(df_data_protection) == pd.DataFrame
    df_data_protection.columns = df_data_protection.columns.str.replace("\n", "")
    return df_data_protection


def dataframe_to_markdown(df: pd.DataFrame, sheet: str = "infosec") -> str:
    markdown_lines = []
    for _, row in df.iterrows():
        levels = row[df.columns[excel_inidces["controlnum"]]].count(".") + 1
        header = "#" * (levels + 1)
        markdown_lines.append(
            f"{header} {row[excel_inidces["controlnum"]]} {row[excel_inidces["controlquestion"]]}"
        )
        dscr = ""
        if (levels > 2) & (sheet == "infosec"):
            # Template String for infosec and prototype
            dscr += f"""\n **{df.columns[excel_inidces["goal"]]}**"""
            dscr += f"""\n {row[excel_inidces["goal"]]}\n"""
            dscr += f"""\n **{df.columns[excel_inidces["requirement_must"]]}**\n"""
            dscr += f"""\n {row[excel_inidces["requirement_must"]]}\n"""
            dscr += f"""\n **{df.columns[excel_inidces["requirement_should"]]}**\n"""
            dscr += f"""\n {row[excel_inidces["requirement_should"]]}\n"""
            dscr += f"""\n **{df.columns[excel_inidces["requirement_high"]]}**\n"""
            dscr += f"""\n {row[excel_inidces["requirement_high"]]}\n"""
            dscr += f"""\n **{df.columns[excel_inidces["requirement_very_high"]]}**\n"""
            dscr += f"""\n {row[excel_inidces["requirement_very_high"]]}\n"""
        if (levels > 2) & (sheet == "prototype"):
            # Template String for data protection
            dscr += f"""\n **{df.columns[excel_inidces["goal"]]}**"""
            dscr += f"""\n {row[excel_inidces["goal"]]}\n"""
            dscr += f"""\n **{df.columns[excel_inidces["requirement_must"]]}**\n"""
            dscr += f"""\n {row[excel_inidces["requirement_must"]]}\n"""
            dscr += f"""\n **{df.columns[excel_inidces["requirement_should"]]}**\n"""
            dscr += f"""\n {row[excel_inidces["requirement_should"]]}\n"""
            dscr += f"""\n **{df.columns[excel_inidces["requirement_high"]]}**\n"""
            dscr += f"""\n {row[excel_inidces["requirement_high"]]}\n"""
        if (levels > 1) & (sheet == "data_protection") & ("5_1" in args.version):
            # Template Strings for data protection
            dscr += f"""\n **{df.columns[excel_inidces["goal"]]}**"""
            dscr += f"""\n {row[excel_inidces["goal"]]}\n"""
        if (levels > 1) & (sheet == "data_protection") & ("6" in args.version):
            dscr += f"""\n **{df.columns[excel_inidces["goal"]]}**"""
            dscr += f"""\n {row[excel_inidces["goal"]]}\n"""
            dscr += f"""\n **{df.columns[excel_inidces["requirement_must"]]}**\n"""
            dscr += f"""\n {row[excel_inidces["requirement_must"]]}\n"""

        markdown_lines.append(dscr)
    markdown_lines.append("\n")
    return "\n".join(markdown_lines)


df = read_infosec(excel_inidces)

basename = (".").join(args.input.split(".")[:-1])
output = f"# {basename}\n\n"

output += dataframe_to_markdown(df)
if args.prototype:
    df_prototype = read_prototype(excel_inidces)
    output += dataframe_to_markdown(df_prototype, sheet="prototype")
if args.data_protection:
    df_data = read_data_protection(excel_inidces)
    output += dataframe_to_markdown(df_data, sheet="data_protection")

# fix formatting issues
# replace all kinds of hyphens with ascii symbols
output = re.sub(r"[‐᠆﹣－⁃−–]+", "-", output)
# replace non breaking space with whitespace
output = re.sub(r"[ ]+", " ", output)
# make items starting with '-' subitems (some are not indented)
output = output.replace("\n-", "\n  -")
# some lines end with a line break, resulting in double line breaks
output = re.sub(r"\n\n\n", "\n\n", output)
# some bullet points start with a " +"
output = re.sub(r" \+", "+", output)
# make all bullet points -, subpoints are indented
output = re.sub(r"\n\+", "\n-", output)

print(output)
with open(args.output, "w") as f:
    f.write(output)
