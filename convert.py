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
    return ord(char) - ord("A")


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
    "goal": c2int("I"),
    "requirement_must": 10,
    "requirement_should": 11,
    "requirement_high": 12,
    "requirement_very_high": 13,
    "documentation": c2int("F"),
    "proof": c2int("G"),
}  # For ISA6_DE_6


match args.version:
    case "6_DE":

        excel_inidces = {
            "skiprows": 1,  # first row is a title
            "sheet_infosec": 4,  # 5th sheet
            "nrows_infosec": 64,
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
            "documentation": c2int("E"),
            "proof": c2int("F"),
        }  # For ISA6_DE_6
    case "5_1_DE":
        pass  # default dict
    case _:
        raise Exception(
            "Sorry, Excel Columns not yet defined in Script. Only version 6_DE and 5_1_DE implemented."
        )


def _read_excel_as_dataframe(
    excel_inidces: dict, sheet, skiprows, nrows
) -> pd.DataFrame:
    df = pd.read_excel(
        args.input,
        skiprows=skiprows,
        sheet_name=sheet,
        nrows=nrows,
        dtype=str,
    )
    assert type(df) == pd.DataFrame
    df.columns = df.columns.str.replace("\n", "")
    return df


def read_infosec(excel_inidces: dict) -> pd.DataFrame:
    df = _read_excel_as_dataframe(
        excel_inidces,
        sheet=excel_inidces["sheet_infosec"],
        skiprows=excel_inidces["skiprows"],
        nrows=excel_inidces["nrows_infosec"],
    )
    return df


def read_prototype(excel_inidces: dict) -> pd.DataFrame:
    df = _read_excel_as_dataframe(
        excel_inidces,
        sheet=excel_inidces["sheet_prototype"],
        skiprows=excel_inidces["skiprows"],
        nrows=excel_inidces["nrows_prototype"],
    )
    return df


def read_data_protection(excel_inidces: dict) -> pd.DataFrame:
    df = _read_excel_as_dataframe(
        excel_inidces,
        skiprows=excel_inidces["skiprows"],
        sheet=excel_inidces["sheet_data_protection"],
        nrows=excel_inidces["nrows_data_protection"],
    )
    return df


def fix_excel_formatting(output: str):
    # fix formatting issues
    # replace all kinds of hyphens with ascii symbols
    output = re.sub(r"[‐᠆﹣－⁃−–]+", "-", output)
    # replace non breaking space with whitespace
    output = re.sub(r"[ ]+", " ", output)
    # # make items starting with '-' subitems (some are not indented)
    output = output.replace("\n-", "\n  -")
    # some lines end with a line break, resulting in double line breaks
    output = re.sub(r"\n\n\n", "\n\n", output)
    # some bullet points start with a " +"
    output = re.sub(r" \+", "+", output)
    # make all bullet points -, subpoints are indented
    output = re.sub(r"\n\+", "\n-", output)
    return output


def dataframe_to_markdown(df: pd.DataFrame, sheet: str = "infosec") -> str:
    markdown_lines = []
    for _, row in df.iterrows():
        print(row[df.columns[excel_inidces["controlnum"]]])
        levels = row[df.columns[excel_inidces["controlnum"]]].count(".") + 1
        header = "#" * (levels + 1)
        markdown_lines.append(
            f"{header} {row[excel_inidces["controlnum"]]} {row[excel_inidces["controlquestion"]]}"
        )
        control_descrition = ""
        implementation = ""
        if (levels > 2) & (sheet == "infosec"):
            # Template String for infosec and prototype
            control_descrition += f"""\n**{df.columns[excel_inidces["goal"]]}**"""
            control_descrition += f"""\n{row[excel_inidces["goal"]]}\n"""
            control_descrition += (
                f"""\n**{df.columns[excel_inidces["requirement_must"]]}**\n"""
            )
            control_descrition += f"""\n{row[excel_inidces["requirement_must"]]}\n"""
            control_descrition += (
                f"""\n**{df.columns[excel_inidces["requirement_should"]]}**\n"""
            )
            control_descrition += f"""\n{row[excel_inidces["requirement_should"]]}\n"""
            control_descrition += (
                f"""\n**{df.columns[excel_inidces["requirement_high"]]}**\n"""
            )
            control_descrition += f"""\n{row[excel_inidces["requirement_high"]]}\n"""
            control_descrition += (
                f"""\n**{df.columns[excel_inidces["requirement_very_high"]]}**\n"""
            )
            control_descrition += (
                f"""\n{row[excel_inidces["requirement_very_high"]]}\n"""
            )
            implementation += (
                f"""\n**{df.columns[excel_inidces["documentation"]]}**\n"""
            )
            implementation += f"""\n{row[excel_inidces["documentation"]]}\n"""
            implementation += f"""\n**{df.columns[excel_inidces["proof"]]}**\n"""
            implementation += f"""\n{row[excel_inidces["proof"]]}\n"""
        if (levels > 2) & (sheet == "prototype"):
            # Template String for data protection
            control_descrition += f"""\n**{df.columns[excel_inidces["goal"]]}**"""
            control_descrition += f"""\n{row[excel_inidces["goal"]]}\n"""
            control_descrition += (
                f"""\n**{df.columns[excel_inidces["requirement_must"]]}**\n"""
            )
            control_descrition += f"""\n{row[excel_inidces["requirement_must"]]}\n"""
            control_descrition += (
                f"""\n**{df.columns[excel_inidces["requirement_should"]]}**\n"""
            )
            control_descrition += f"""\n{row[excel_inidces["requirement_should"]]}\n"""
            control_descrition += (
                f"""\n**{df.columns[excel_inidces["requirement_high"]]}**\n"""
            )
            control_descrition += f"""\n{row[excel_inidces["requirement_high"]]}\n"""
        if (levels > 1) & (sheet == "data_protection") & ("5_1" in args.version):
            # Template Strings for data protection
            control_descrition += f"""\n**{df.columns[excel_inidces["goal"]]}**"""
            control_descrition += f"""\n{row[excel_inidces["goal"]]}\n"""
        if (levels > 1) & (sheet == "data_protection") & ("6" in args.version):
            control_descrition += f"""\n**{df.columns[excel_inidces["goal"]]}**"""
            control_descrition += f"""\n{row[excel_inidces["goal"]]}\n"""
            control_descrition += (
                f"""\n**{df.columns[excel_inidces["requirement_must"]]}**\n"""
            )
            control_descrition += f"""\n{row[excel_inidces["requirement_must"]]}\n"""

        control_descrition = fix_excel_formatting(control_descrition)
        markdown_lines.append(control_descrition + implementation)
    markdown_lines.append("\n")
    return "\n".join(markdown_lines)


df = read_infosec(excel_inidces)

# basename = (".").join(args.input.split(".")[:-1])
# output = f"# {basename}\n\n"
output = ""

output += dataframe_to_markdown(df)
if args.prototype:
    df_prototype = read_prototype(excel_inidces)
    output += dataframe_to_markdown(df_prototype, sheet="prototype")
if args.data_protection:
    df_data = read_data_protection(excel_inidces)
    output += dataframe_to_markdown(df_data, sheet="data_protection")


print(output)
with open(args.output, "w") as f:
    f.write(output)
