import docx
import os
import shutil

class Preprocessor:

    def delete_before_heading1(self, doc):
        """
        Deletes all pre-body content in the document.
        """
        body = doc.element.body
        elements = list(body)  # Obtain all elements in the document (paragraphs and tables)

        for idx, element in enumerate(elements):
            if element.tag.endswith('p'): # Check if it's a paragraph element
                para = docx.text.paragraph.Paragraph(element, doc)
                if para.style and para.style.name == "Heading 1":
                    # Delete all elements before the first "Heading 1"
                    for e in elements[:idx]:
                        body.remove(e)
                    print("All pre-body content has been deleted.")
                    return
        print("Unable to find Heading 1 in the document, no content deleted.")

    def delete_headers_footers(self, doc):
        """
        Delete all headers and footers from the document.
        """
        for section in doc.sections:
            # Delete headers and footers
            section.header.is_linked_to_previous = False
            section.footer.is_linked_to_previous = False

            for paragraph in section.header.paragraphs:
                p = paragraph._element
                p.getparent().remove(p)
            for paragraph in section.footer.paragraphs:
                p = paragraph._element
                p.getparent().remove(p)

        print("All headers and footers have been deleted.")

    def remove_watermark(self, doc):
        """
        Try to remove watermark from the document (based on XML structure).
        Notice: This method may fail if the document structure changes.
        """
        try:
            # Access the XML content of the document and delete the watermark node
            for element in doc.element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}header'):
                for watermark in element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pict'):
                    element.remove(watermark)
            print("Watermark has been removed, if any.")
        except Exception as e:
            print(f"An error occurred while removing watermark: {e}")

    def process_word_file(self, input_path, output_path):
        """
        Process a Word file: delete all content before the first "Heading 1", delete headers, footers, and watermark.
        """
        # Create a backup path for the output file
        backup_path = f"{output_path}.backup.docx"
        try:
            shutil.copyfile(input_path, backup_path)
            print(f"A backup of the original file has been created at {backup_path}")

        except Exception as e:
            print(f"An error occurred while creating the backup: {e}")
            return

        try:
            doc = docx.Document(input_path)
            self.delete_before_heading1(doc)
            self.delete_headers_footers(doc)
            self.remove_watermark(doc)
            # Save the processed document to the output path
            doc.save(output_path)
            print(f"Doc preprocessing completed. The processed document is saved at {output_path}")

        except Exception as e:
            print(f"An error occurred during document preprocessing: {e}")
