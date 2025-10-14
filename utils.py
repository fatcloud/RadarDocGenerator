import pathlib

def concatenate_docx(doc_paths, output_docx_path):
    from docxcompose.composer import Composer
    from docx import Document as Document_compose

    # Create a new document to act as the master, then copy styles and page layout from the template.
    master = Document_compose()
    template = Document_compose('templates/baseline.docx')

    # Copy styles from template to master
    master_styles_element = master.styles.element
    template_styles_element = template.styles.element
    del master_styles_element[:]
    for style_element in template_styles_element:
        master_styles_element.append(style_element)

    # Copy section properties from template to master
    for s_template, s_master in zip(template.sections, master.sections):
        s_master.top_margin = s_template.top_margin
        s_master.bottom_margin = s_template.bottom_margin
        s_master.left_margin = s_template.left_margin
        s_master.right_margin = s_template.right_margin
        s_master.page_height = s_template.page_height
        s_master.page_width = s_template.page_width
        s_master.orientation = s_template.orientation

    composer = Composer(master)

    for doc_path in doc_paths:
        composer.append(Document_compose(doc_path))

    pathlib.Path(output_docx_path).parent.mkdir(parents=True, exist_ok=True)
    composer.save(output_docx_path)