from docx import Document



files = ['D:\\python3\\lubbock1019\\2019012484\\Letter.docx', 'D:\\python3\\lubbock1019\\2019012483\\Letter.docx', 'D:\\python3\\lubbock1019\\2019012482\\Letter.docx', 'D:\\python3\\lubbock1019\\2019012481\\Letter.docx', 'D:\\python3\\lubbock1019\\2019012477\\Letter.docx']


def combine_word_documents(files):
    merged_document = Document()

    for index, file in enumerate(files):
        sub_doc = Document(file)

        # Don't add a page break if you've reached the last file.
        if index < len(files)-1:
           sub_doc.add_page_break()

        for element in sub_doc.element.body:
            merged_document.element.body.append(element)

    merged_document.save('merged.docx')

combine_word_documents(files)