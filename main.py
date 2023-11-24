import os

import piexif as piexif
import xlrd
import xlsxwriter
from PIL import Image


class ImageResizer:
    def __init__(self, input_path, output_path):
        self.input_path = input_path
        self.output_path = output_path
        self.default_pairs = [
            (200, 200),
        ]

    def resize_and_convert_image(self, max_width=150, max_height=150):
        # Open the input image
        img = Image.open(self.input_path)
        width, height = img.size
        scale = min(max_width / width, max_height / height)
        new_width = int(width * scale)
        new_height = int(height * scale)
        # Resize the image to the specified dimensions
        img = img.resize((new_width, new_height))

        # Convert the image to JPEG format
        # img = img.convert('RGB')

        # Construct the output filename with the specified prefix
        output_path = self.output_path

        # Remove the EXIF metadata from the image
        try:
            exif_dict = piexif.load(img.info['exif'])
            exif_bytes = piexif.dump(exif_dict)
            img.save(output_path, exif=exif_bytes)
        except (KeyError, ValueError, TypeError):
            img.save(output_path)

    def resize_and_convert_default_pairs(self):
        for width, height in self.default_pairs:
            self.resize_and_convert_image(width, height)


class Images():
    def __init__(self, path, codename):
        self.path = path
        self.codename = codename


class Processing():
    BASE_DIR = "."
    DATA_DIR = os.path.join(BASE_DIR, "images")
    XLSX_FNAME = os.path.join(DATA_DIR, "ΚΑΤΑΛΟΓΟΣ.xls")

    def get_sheet(self):
        workbook = xlrd.open_workbook(self.XLSX_FNAME)
        return workbook.sheet_by_index(0)

    def get_images(self, ):
        images = []

        for file in os.listdir(self.DATA_DIR):
            suffixes = ["png", "jpg", "jpeg"]
            ending = file.split(".")[-1]

            starting = file.split(f".{ending}")[0]
            path = os.path.join(self.DATA_DIR, file)
            if ending in suffixes:
                resizer = ImageResizer(path, path)
                resizer.resize_and_convert_default_pairs()
                img = Images(codename=starting, path=path)
                images.append(img)
            else:
                raise Exception(f"Παρε με τηλεφωνο και πες μου οτι δεν υποστηριζει καταληξη {ending}")
        return images

    def get_col_names(self, sheet):
        names = []
        for elem in range(sheet.ncols):
            names.append(sheet.col_values(elem)[0])
        return names

    def write_new_colums(self, sheet, old_sheet):
        new_column_data = self.get_col_names(old_sheet) + ["images", ]
        for row_num, value in enumerate(new_column_data):
            sheet.write(0, row_num, value)

    def copy_old_sheet(self, sheet, old_sheet):
        for row_index in range(old_sheet.nrows):
            for col_index in range(old_sheet.ncols):
                sheet.write(row_index, col_index, old_sheet.cell_value(row_index, col_index))

    def create_new_sheet(self, old_sheet):
        images = self.get_images()
        new_workbook = xlsxwriter.Workbook("new.xls")

        sheet = new_workbook.add_worksheet('Sheet1')
        sheet.set_default_row(170)
        sheet.set_column('D:D', 40)

        self.write_new_colums(sheet, old_sheet)
        self.copy_old_sheet(sheet, old_sheet)

        for idx, col_value in enumerate(old_sheet.col_values(1)):
            for image in images:
                if image.codename == col_value:
                    data = {
                        'x_offset': 70,
                        'y_offset': 1,
                        'x_scale': 1,
                        'y_scale': 1,
                        'object_position': 0,
                    }
                    sheet.insert_image(idx, 3, image.path, options=data)

        sheet.autofit()
        return new_workbook, sheet

    def run(self):
        sheet = self.get_sheet()
        new_workbook, new_sheet = self.create_new_sheet(sheet)
        new_workbook.close()


def main():
    processor = Processing()
    processor.run()


main()
