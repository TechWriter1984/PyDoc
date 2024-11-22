import docx

class FileChecker:
    def __init__(self, doc_path):
        self.doc = docx.Document(doc_path)

    def print_paragraph_styles_and_xml(self):
        """Prints the style name and XML content for each paragraph in the document."""
        for i, para in enumerate(self.doc.paragraphs):
            print(f"Paragraph {i+1} - Style: {para.style.name if para.style else 'No Style'}")
            print(f"Text: {para.text.strip()}")
            for run in para.runs:
                # 打印每个运行项的 XML 内容
                print(f"Run XML: {run._element.xml}")
            print("-" * 50)

# 测试代码
if __name__ == "__main__":
    # 替换为你的文件路径
    file_path = "D:\doc_marco_dev\data\input\(HS-SC-CT10W0C3) Hanshow Smart Cart Installation Manual V1.0.1.docx"
    checker = FileChecker(file_path)
    checker.print_paragraph_styles_and_xml()
