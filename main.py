import pathlib
from docx import Document
from docx.shared import Pt
from docx2pdf import convert

def main():
    print("This program currently only works with .docx files. If you have a Word compatibility file \n"
          "please save it as a .docx file instead.\n")

    name = str(input("What's your name? "))
    # Get the current working directory
    current_dir = str(pathlib.Path(__file__).parent.resolve())

    # Append name to file path for the target file path
    docx_path_path = get_target_document_path(current_dir)

    # Get job title and company name applying for
    job_title, company_name = get_job_title_and_company_name()

    # Open the document
    document = open_document(docx_path_path)

    # Ask user for font and sizing
    try:
        style = get_current_style(document)
    except Exception:
        # style is None is shut PyCharm up
        style = None
        exit("Style font or sizing is incorrect.")

    # Replace all instances of "jobtitle" and "thecompany" with corresponding inputs
    replace_text("jobtitle", job_title, document, style)
    replace_text("thecompany", company_name, document, style)

    # Save the document
    document.save(name + 'Cover Letter ' + company_name + '.docx')
    # Convert to PDF
    convert(name + 'Cover Letter ' + company_name + '.docx', name + ' Cover Letter ' + company_name + '.pdf')
    print("\nPDF version dumped to Documents folder.")


def get_job_title_and_company_name():
    job_title = str(input("Job title: "))
    company_name = str(input("Company name: "))

    return job_title, company_name


def get_current_style(document):
    # Get font and font size from user
    style = document.styles['Normal']
    font = style.font
    font.name = str(input("Please enter a font name to use: "))
    font.size = Pt(int(input("What font size: ")))

    return style


def replace_text(search, replace, document, style):
    # Find matching words
    for paragraph in document.paragraphs:
        if search in paragraph.text:
            print("Found match for " + search)

            # Replace words and apply style as specified.
            new_para = paragraph.text.replace(search, replace)
            paragraph.text = new_para
            paragraph.style = style


def open_document(path_to_file):
    try:
        document = Document(path_to_file)
    except Exception:
        exit("Failed to open document.")
        # Putting this here after the exit so PyCharm shuts up
        return

    print("Successfully opened document.")

    return document


# Ask for file name and return correct file path
def get_target_document_path(current_dir):
    docx_file_name = str(input("Please type the name of the Word (.docx) file you want to edit: "))

    docx_file_path = current_dir + "\\" + docx_file_name + ".docx"

    return docx_file_path


if __name__ == "__main__":
    main()
