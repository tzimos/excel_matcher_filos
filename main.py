import argparse
import os
import pandas
import xlsxwriter


def get_parser():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        '--excel-file',
        "-e",
        required=True,
        help='The full path of the target excel file.'
    )
    parser.add_argument(
        '--images-dir',
        "-i",
        required=True,
        help='The full path of the directory of images.'
    )
    parser.add_argument(
        '--output-dir',
        "-o",
        default=r"C:\Users\asima\Desktop",
        help='The full path of the directory of images.'
    )
    return parser


def validate_exists(file_path):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f'File path not found: {file_path}')


def get_image_paths(images_dir):
    return [
        os.path.join(images_dir, f) for f in os.listdir(images_dir)
    ]


def main():
    parser = get_parser()
    args = parser.parse_args()
    validate_exists(args.excel_file)
    validate_exists(args.images_dir)
    df = pandas.read_excel(args.excel_file)
    image_paths = get_image_paths(args.images_dir)
    df['images'] = image_paths
    output_file = os.path.join(args.output_dir, 'arxidia.xlsx')

    book = xlsxwriter.Workbook(output_file)
    worksheet = book.add_worksheet()

    for row, col in df.iterrows():
        worksheet.write(row, 0, col['ΚΩΔΙΚΟΣ'])
        worksheet.write(row, 1, col['ΠΕΡΙΓΡΑΦΗ'])
        worksheet.insert_image(row, 2, filename=col['images'], options={'x_scale': 0.04, 'y_scale': 0.04})
    worksheet.set_default_row(40)
    worksheet.set_column(2, 1, 40)

    book.close()


main()
