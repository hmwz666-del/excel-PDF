import os
import tempfile
import unittest
from types import SimpleNamespace
from unittest.mock import patch

from converter import ConversionResult, ExcelConverter


class FakeWorkbook:
    def __init__(self):
        self.closed = False

    def ExportAsFixedFormat(self, Type, Filename, **kwargs):
        with open(Filename, "wb") as temp_pdf:
            temp_pdf.write(b"%PDF-1.4\n")

    def Close(self, SaveChanges=False):
        self.closed = True


class FakeWorkbooks:
    def __init__(self, workbook):
        self.workbook = workbook

    def Open(self, *args, **kwargs):
        return self.workbook


class FakeExcelApp:
    def __init__(self, workbook):
        self.Workbooks = FakeWorkbooks(workbook)


class FakeContentStream:
    def __init__(self, data):
        self._data = data

    def get_data(self):
        return self._data


class FakePdfObject(dict):
    def get_object(self):
        return self


class FakePage(dict):
    def __init__(self, text="", xobjects=None, contents=None):
        super().__init__()
        self._text = text
        self._contents = contents

        if xobjects is not None:
            self["/Resources"] = {"/XObject": xobjects}

    def extract_text(self):
        return self._text

    def get_contents(self):
        return self._contents


class ConverterTests(unittest.TestCase):
    def test_convert_file_removes_plaintext_temp_pdf_when_copy_fails(self):
        converter = ExcelConverter()
        converter._initialized = True

        workbook = FakeWorkbook()
        converter.excel_app = FakeExcelApp(workbook)

        with tempfile.TemporaryDirectory() as temp_dir, tempfile.TemporaryDirectory() as output_dir:
            source_path = os.path.join(temp_dir, "source.xlsx")
            temp_pdf_path = os.path.join(temp_dir, "excel_to_pdf_temp_fixed.pdf")

            with patch("converter.RETRY_COUNT", 0), patch.object(
                converter,
                "_prepare_workbook",
                return_value=None,
            ), patch.object(
                converter,
                "_remove_last_blank_page",
                return_value=False,
            ), patch(
                "converter.tempfile.gettempdir",
                return_value=temp_dir,
            ), patch(
                "converter.uuid.uuid4",
                return_value=SimpleNamespace(hex="fixed"),
            ), patch(
                "converter.shutil.copy2",
                side_effect=OSError("copy failed"),
            ):
                result = converter.convert_file(source_path, output_dir)

            self.assertEqual(result.status, ConversionResult.FAILED)
            self.assertFalse(os.path.exists(temp_pdf_path))
            self.assertTrue(workbook.closed)

    def test_form_xobject_page_counts_as_meaningful_content(self):
        converter = ExcelConverter()
        page = FakePage(
            xobjects={"/Fm0": FakePdfObject({"/Subtype": "/Form"})},
            contents=FakeContentStream(b""),
        )

        self.assertTrue(converter._page_has_meaningful_content(page))

    def test_empty_page_without_text_or_drawing_is_not_meaningful(self):
        converter = ExcelConverter()
        page = FakePage(text="", xobjects={}, contents=FakeContentStream(b""))

        self.assertFalse(converter._page_has_meaningful_content(page))


if __name__ == "__main__":
    unittest.main()
