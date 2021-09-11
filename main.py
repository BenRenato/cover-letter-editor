import os
import pathlib
import docx
from docx import Document


def main():
    print("This program currently only works with .docx files. If you have a Word compatibility file \n"
          "please save it as a .docx file instead.")

    # Get the current working directory
    current_dir = str(pathlib.Path(__file__).parent.resolve())

    docx_path_path = get_target_document_path(current_dir)

    document = open_document(docx_path_path)

    for p in document.paragraphs:
        print(p.text)


def open_document(path_to_file):
    try:
        document = Document(path_to_file)
    except Exception:
        exit("Failed to open document.")
        # Putting this here after the exit so PyCharm shuts up
        return

    return document


def get_target_document_path(current_dir):
    docx_file_name = str(input("Please type the name of the Word (.docx) file you want to edit: "))

    docx_file_path = current_dir + "\\" + docx_file_name + ".docx"

    return docx_file_path


if __name__ == "__main__":
    main()
