import mtranslate
import win32com.client as win32

class ShapeTranslator:
    def __init__(self, file_path):
        self.file_path = file_path

    def rename_shapes_within_group(self, group_shape, worksheet_name):
        """Renames the shapes within a group shape."""
        for shape in group_shape.GroupItems:
            original_name = shape.Name
            translated_name = mtranslate.translate(original_name, "en", "auto")
            shape.Name = translated_name
            print('worksheet', worksheet_name, 'original_name', original_name, 'translated_name', translated_name)

    def rename_shapes_in_worksheet(self, worksheet):
        """Renames the shapes within a worksheet, including shapes within groups."""
        shape_count = worksheet.Shapes.Count
        for i in range(1, shape_count + 1):
            shape = worksheet.Shapes.Item(i)
            if shape.Type == 6:
                self.rename_shapes_within_group(shape, worksheet.Name)
            else:
                original_name = shape.Name
                translated_name = mtranslate.translate(original_name, "en", "auto")
                shape.Name = translated_name
                print('worksheet', worksheet.Name, 'original_name', original_name, 'translated_name', translated_name)

    def translate_shape_names(self):
        """Translates the names of shapes within all sheets of the workbook from Japanese to English."""
        excel_app = win32.Dispatch("Excel.Application")
        workbook = excel_app.Workbooks.Open(self.file_path)
        for worksheet in workbook.Sheets:
            self.rename_shapes_in_worksheet(worksheet)
        workbook.Save()
        workbook.Close()
        excel_app.Quit()

excel_file = "C:/MySpace/Projects/learning/DeformationCheck/data/HA24K230405(JHV723EJ)径寸法データNew4.xlsm"

translator = ShapeTranslator(excel_file)
translator.translate_shape_names()
