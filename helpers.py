import yaml
from docx import Document

def load_yaml_config(yaml_path):
    """Charge le fichier de configuration YAML."""
    with open(yaml_path, "r", encoding="utf-8") as file:
        return yaml.safe_load(file)

def replace_text_in_docx(doc_path, yaml_path, output_path):
    """Remplace les valeurs dans un document Word selon un fichier YAML."""
    # Charger le fichier YAML
    replacements = load_yaml_config(yaml_path)

    # Charger le document Word
    doc = Document(doc_path)

    # Parcourir tous les paragraphes et remplacer les valeurs
    for paragraph in doc.paragraphs:
        for key, value in replacements.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)

    # Sauvegarder le document modifié
    doc.save(output_path)
    print(f"Document modifié enregistré sous : {output_path}")

# Exemple d'utilisation
doc_path = "template.docx"       # Chemin du document d'origine
yaml_path = "config.yaml"        # Chemin du fichier de configuration YAML
output_path = "output.docx"      # Chemin du document modifié

replace_text_in_docx(doc_path, yaml_path, output_path)




