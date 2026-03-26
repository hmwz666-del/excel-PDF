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


class FakeFont:
    def __init__(self, size=11):
        self.Size = size


class FakeMergeDimension:
    def __init__(self, count):
        self.Count = count


class FakeMergeArea:
    def __init__(self, column, columns_count, width):
        self.Column = column
        self.Columns = FakeMergeDimension(columns_count)
        self.Width = width


class FakeGridCell:
    def __init__(
        self,
        row,
        column,
        value=None,
        text=None,
        width=30,
        wrap_text=False,
        merge_area=None,
        font_size=11,
    ):
        self.Row = row
        self.Column = column
        self.Value = value
        self.Text = text if text is not None else ("" if value is None else str(value))
        self.Width = width
        self.WrapText = wrap_text
        self.MergeArea = merge_area
        self.MergeCells = merge_area is not None
        self.Font = FakeFont(font_size)


class FakeColumnCollection:
    def __init__(self, count):
        self.Count = count


class FakeSheet:
    def __init__(self, cells, column_count=20):
        self._cells = cells
        self.Columns = FakeColumnCollection(column_count)

    def Cells(self, row, column):
        return self._cells[(row, column)]


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

    def test_get_cell_visual_right_col_uses_merge_area_boundary(self):
        converter = ExcelConverter()
        merge_area = FakeMergeArea(column=5, columns_count=3, width=120)
        cell = FakeGridCell(
            row=2,
            column=5,
            value="MINISO26030700185",
            width=40,
            merge_area=merge_area,
        )
        sheet = FakeSheet({(2, 5): cell})

        self.assertEqual(converter._get_cell_visual_right_col(sheet, cell, cell.Value), 7)

    def test_get_cell_visual_right_col_expands_into_blank_columns_for_overflow_text(self):
        converter = ExcelConverter()
        cell = FakeGridCell(
            row=1,
            column=5,
            value="MINISO260307001853",
            width=28,
            font_size=11,
        )
        next_cell_1 = FakeGridCell(row=1, column=6, value=None, width=32)
        next_cell_2 = FakeGridCell(row=1, column=7, value=None, width=32)
        next_cell_3 = FakeGridCell(row=1, column=8, value=None, width=32)
        next_cell_4 = FakeGridCell(row=1, column=9, value=None, width=32)
        sheet = FakeSheet(
            {
                (1, 5): cell,
                (1, 6): next_cell_1,
                (1, 7): next_cell_2,
                (1, 8): next_cell_3,
                (1, 9): next_cell_4,
            }
        )

        self.assertGreaterEqual(
            converter._get_cell_visual_right_col(sheet, cell, cell.Value),
            6,
        )


if __name__ == "__main__":
    unittest.main()
