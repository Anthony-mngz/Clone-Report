import yaml
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Inches
import pandas as pd


def load_yaml_config(yaml_path):
    """Charge le fichier de configuration YAML."""
    with open(yaml_path, "r", encoding="utf-8") as file:
        return yaml.safe_load(file)


def apply_format(run, format_config):
    if "font" in format_config:
        run.font.name = format_config["font"]
    if "size" in format_config:
        run.font.size = Pt(format_config["size"])
    if "bold" in format_config:
        run.bold = format_config["bold"]


def replace_text(paragraph, replacements):
    for key, value in replacements.items():
        if isinstance(value, (int, float)) and "rounding" in replacements:
            value = round(value, replacements["rounding"])
        if key in paragraph.text:
            paragraph.text = paragraph.text.replace(key, str(value))


def insert_image(doc, image_path, size):
    if size:
        width, height = size
        doc.add_picture(image_path, width=Inches(width), height=Inches(height))
    else:
        doc.add_picture(image_path)


def insert_table(doc, data):
    table = doc.add_table(rows=len(data) + 1, cols=len(data[0]))
    for col_idx, header in enumerate(data[0].keys()):
        table.cell(0, col_idx).text = header
    for row_idx, row in enumerate(data):
        for col_idx, cell in enumerate(row.values()):
            table.cell(row_idx + 1, col_idx).text = str(cell)


def replace_text_in_docx(doc_path, yaml_path, output_path):
    config = load_yaml_config(yaml_path)
    doc = Document(doc_path)

    for paragraph in doc.paragraphs:
        if "text" in config:
            replace_text(paragraph, config["text"])
            if "format" in config:
                for run in paragraph.runs:
                    apply_format(run, config["format"])

    if "image" in config:
        insert_image(doc, config["image"]["path"], config["image"].get("size"))

    if "array" in config:
        insert_table(doc, config["array"])

    if "array_from_excel" in config:
        df = pd.read_excel(config["array_from_excel"]["path"])
        insert_table(doc, df.to_dict(orient="records"))

    doc.save(output_path)
    print(f"Document modifié enregistré sous : {output_path}")


# Exemple d'utilisation
doc_path = "template.docx"
yaml_path = "config.yaml"
output_path = "output.docx"
replace_text_in_docx(doc_path, yaml_path, output_path)

