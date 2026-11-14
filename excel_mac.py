"""Expose the Excel operations shared with the Windows workbook writer on macOS."""
from contextlib import contextmanager
from pathlib import Path


def rgb(bgr):
    return bgr & 255, (bgr >> 8) & 255, (bgr >> 16) & 255


class Font:
    def __init__(self, font):
        self.font = font

    @property
    def Bold(self):
        return self.font.bold

    @Bold.setter
    def Bold(self, value):
        self.font.bold = value

    @property
    def Color(self):
        return self.font.color

    @Color.setter
    def Color(self, value):
        self.font.color = rgb(value)


class Range:
    def __init__(self, cell_range):
        self.range = cell_range

    # xlwings exposes these operations as public properties on a Range.
    Value2 = property(lambda self: self.range.value, lambda self, value: setattr(self.range, "value", value))
    Formula = property(lambda self: self.range.formula, lambda self, value: setattr(self.range, "formula", value))
    ColumnWidth = property(lambda self: self.range.column_width, lambda self, value: setattr(self.range, "column_width", value))
    RowHeight = property(lambda self: self.range.row_height, lambda self, value: setattr(self.range, "row_height", value))
    WrapText = property(lambda self: self.range.wrap_text, lambda self, value: setattr(self.range, "wrap_text", value))
    NumberFormat = property(lambda self: self.range.number_format, lambda self, value: setattr(self.range, "number_format", value))

    @property
    def Interior(self):
        return self

    @property
    def Color(self):
        return self.range.color

    @Color.setter
    def Color(self, value):
        self.range.color = rgb(value)

    @property
    def Font(self):
        return Font(self.range.font)

    @property
    def HorizontalAlignment(self):
        return self.range.api.horizontal_alignment.get()

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, value):
        from appscript import k
        alignment = {-4108: k.horizontal_align_center, -4131: k.horizontal_align_left}
        self.range.api.horizontal_alignment.set(alignment[value])

    def Borders(self, edge):
        from appscript import k
        if edge != 9:
            raise ValueError(f"Unsupported border edge: {edge}")
        return Border(self.range.api.get_border(which_border=k.edge_bottom))

    def Merge(self):
        self.range.merge()

    def ClearContents(self):
        self.range.clear_contents()


class Border:
    def __init__(self, border):
        self.border = border

    @property
    def LineStyle(self):
        return self.border.line_style.get()

    @LineStyle.setter
    def LineStyle(self, value):
        from appscript import k
        if value != 1:
            raise ValueError(f"Unsupported line style: {value}")
        self.border.line_style.set(k.continuous)


class Sheet:
    def __init__(self, sheet):
        self.sheet = sheet

    Name = property(lambda self: self.sheet.name, lambda self, name: setattr(self.sheet, "name", name))

    def Cells(self, row, column):
        return Range(self.sheet.range((row, column)))

    def Range(self, address):
        return Range(self.sheet.range(address))

    Columns = Range

    def Delete(self):
        self.sheet.delete()

    def Copy(self, *, After):
        return Sheet(self.sheet.copy(after=After.sheet))

    def Calculate(self):
        self.sheet.book.app.calculate()


class Worksheets:
    def __init__(self, sheets):
        self.sheets = sheets

    @property
    def Count(self):
        return self.sheets.count

    def __iter__(self):
        return (Sheet(sheet) for sheet in self.sheets)

    def __call__(self, key):
        return Sheet(self.sheets[key - 1 if isinstance(key, int) else key])

    def Add(self, *, After):
        return Sheet(self.sheets.add(after=After.sheet))


class Workbook:
    def __init__(self, book):
        self.book = book
        self.Worksheets = Worksheets(book.sheets)

    @property
    def ReadOnly(self):
        return bool(self.book.api.read_only.get())

    @property
    def ActiveSheet(self):
        return Sheet(self.book.sheets.active)

    def Save(self):
        self.book.save()

    def Close(self, *, SaveChanges):
        if SaveChanges:
            self.book.save()
        self.book.close()


class Excel:
    def __init__(self, app):
        self.app = app
        self.Workbooks = self

    def Open(self, path, *, UpdateLinks, ReadOnly):
        return Workbook(self.app.books.open(path, update_links=bool(UpdateLinks), read_only=ReadOnly))


@contextmanager
def session():
    import xlwings as xw

    if len(xw.apps):
        raise RuntimeError("Quit Microsoft Excel with Command+Q before running WASDE.")
    app = xw.App(visible=False, add_book=False)
    try:
        app.display_alerts = False
        yield Excel(app)
    finally:
        app.quit()


def check_install():
    """Save and reopen a disposable book to check Excel and Automation permission."""
    from tempfile import TemporaryDirectory

    with TemporaryDirectory(prefix="wasde-excel-") as folder, session() as excel:
        path = Path(folder) / "WASDE Test.xlsx"
        book = excel.app.books.add()
        sheet = Sheet(book.sheets[0])
        copy = sheet.Copy(After=sheet)
        copy.Name = "WASDE Test"
        copy.Cells(1, 1).Formula = "=2+3"
        copy.Range("B2:C2").Merge()
        copy.Range("B2:C2").Font.Bold = True
        copy.Range("B2:C2").Borders(9).LineStyle = 1
        copy.Range("B2:C2").HorizontalAlignment = -4108
        copy.Calculate()
        book.save(str(path))
        book.close()
        book = excel.Open(str(path), UpdateLinks=False, ReadOnly=True)
        try:
            if book.Worksheets("WASDE Test").Cells(1, 1).Value2 != 5:
                raise RuntimeError("Excel did not save or calculate the test formula")
        finally:
            book.Close(SaveChanges=False)
    print("PASS: Excel for Mac created, calculated, saved and reopened a test workbook; five workbook paths found")
