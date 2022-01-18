import pathlib

def concatenate_docx(doc_paths, output_docx_path):
    from docxcompose.composer import Composer
    from docx import Document as Document_compose

    master = Document_compose(doc_paths[0])
    composer = Composer(master)

    for doc_path in doc_paths[1:]:
        partial_docx = Document_compose(doc_path)
        composer.append(partial_docx)

    pathlib.Path(output_docx_path).parent.mkdir(parents=True, exist_ok=True)
    composer.save(output_docx_path)