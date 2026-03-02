# docx_parser.py
import zipfile
import xml.etree.ElementTree as ET


class DOCXParser:
    """Класс для парсинга DOCX файлов"""

    @staticmethod
    def extract_text_from_docx(docx_path: str) -> str:
        """Извлекает текст из DOCX файла"""
        try:
            with zipfile.ZipFile(docx_path) as docx:
                xml_content = docx.read('word/document.xml')
                root = ET.fromstring(xml_content)

                namespaces = {
                    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                }

                paragraphs = root.findall('.//w:p', namespaces)
                text_lines = []

                for para in paragraphs:
                    texts = para.findall('.//w:t', namespaces)
                    para_text = ''.join([text.text if text.text else '' for text in texts])

                    if para_text.strip():
                        text_lines.append(para_text)

                full_text = '\n'.join(text_lines)
                return full_text

        except Exception as e:
            raise Exception(f"Ошибка чтения DOCX файла: {str(e)}")