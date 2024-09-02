<!-- ltex: language=en -->
# Convert Tisax Controls to Markdown (Or anything else)

The Information Security Standard TISAX (VDA ISA) is distributed as a Microsoft Excel file,
see [portal.enx.com/en-us/TISAX/downloads/](https://portal.enx.com/en-us/TISAX/downloads/).
For Reports and a Statement of Applicability you might need the controls as plain text.

I assume Most people manage their Information Security with office Macros and don't need this 🤯.

This script converts the Controls and Requirements to Markdown, you could use [pandoc](https://www.pandoc.org/) to convert them to anything else.

## Usage

```
pip install -r requirements.txt # installs pandas
python3 convert.py --input VDA_ISA_5_1_DE.xlsx --output VDA_ISA5_1_DE.md --version 5_1_DE
```
So far I only configured this for Version 6_DE and 5_1_DE, but it is easy to update.
You could convert the Markdown file to Word with `pandoc VDA_ISA5_1_DE.md -o VDA_ISA5_1_DE.docx`

## Todo

- [ ] Check Settings (Sheet Number, Columns) for the english ISA Versions (maybe they are the same as German 💁)
