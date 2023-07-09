import win32com.client
import PIL
from PIL import ImageGrab, Image
import os
import sys
from pptx import Presentation
from pptx.util import Inches, Cm
from utilities import variables2
import ctypes
from time import sleep
import shutil

class DownloadImage:
    def __init__(self, data):
        """Initialize the class with the paths of the image, input files, and output file."""
        self.image_path = data["image_path"]
        self.ip_file_1 = data["ip_file_1"]
        self.ip_file_2 = data["ip_file_2"]
        self.op_file = data["op_file"]

    def extract_images_from_excel(self):
        """This function extracts a graph from the input excel file and saves it into the specified PNG image path (overwrites the given PNG image)"""
        print("Process input file 1 : " + self.ip_file_1)
        self.save_excel_chart_as_jpg(self.ip_file_1)
        print("Process input file 2 : " + self.ip_file_2)
        self.save_excel_chart_as_jpg(self.ip_file_2)

    def copy_downloaded_images_to_ppt(self):
        """This function inserts downloaded imaged into respective ppt slide with reference to definition available in varaiables_dict"""
        for img_file_name, val in variables2.image_filename_slide_dict.items():
            slide_idx = val["slide_index"]
            print("Inserting " + img_file_name + " into slide " + str(slide_idx) + " of file " + str(self.op_file))
            image_file = str(self.image_path + img_file_name)
            dmns = variables2.dimensions_dict[val["image_type"]]
            self.copy_image(image_file, slide_idx, dmns["image_height"], dmns["image_width"],
                            dmns["horizontal_pos"], dmns["vertical_pos"])


    def save_excel_chart_as_jpg(self, filename):
        """This function to extact shapes and charts from excel and save in Images folder"""
        print("Processing the excel charts and saving it to the images folder in the current directory")
        self.killexcel()
        # Open the excel application using win32com
        o = win32com.client.Dispatch("Excel.Application")
        # Disable alerts and visibility to the user
        o.Visible = 0
        o.DisplayAlerts = 0
        # Open workbook
        wb = o.Workbooks.Open(filename, None, False)
        sleep(5)
        table_reference_dict = {"Summary": variables2.summary_reference_dict,
                                "Graph":  variables2.judgment_reference_dict}

        for m, sheet in enumerate(o.Sheets):
            sheet.Cells.ClearComments()
            if sheet.Name in ["Summary", "Graph"]:
                for table, stages in table_reference_dict[sheet.Name].items():
                    for stage, range_rerence in stages.items():
                        new_img_path = f"{self.image_path}{table}_{stage}.png"
                        copy_range = sheet.Range(range_rerence)
                        print('new_img_path', new_img_path)
                        try_copy_image = True
                        wait_time = 0.5
                        while try_copy_image:
                            try:
                                self.clear_clip_board()
                                copy_range.Copy()
                                sleep(wait_time)
                                image = ImageGrab.grabclipboard().convert('RGB')
                                image.save(new_img_path)
                                try_copy_image = False
                            except:
                                wait_time += 1
                                print('error - retrying, wait_time = ', wait_time)

            for n, shape in enumerate(sheet.Shapes):
                new_img_path = f"{self.image_path}{sheet.Name}_{n}"
                new_img_path = new_img_path.replace(" ", "").replace(".", "") + ".png"
                if os.path.exists(new_img_path):
                    new_img_path = f"{self.image_path}{sheet.Name}_{n}{n}"
                    new_img_path = new_img_path.replace(" ", "").replace(".", "") + ".png"
                print('shape', shape.Name, 'new_img_path', new_img_path)
                try_copy_image = True
                wait_time = 0.5
                while try_copy_image:
                    try:
                        self.clear_clip_board()
                        shape.Copy()
                        sleep(wait_time)
                        image = ImageGrab.grabclipboard().convert('RGB')
                        image.save(new_img_path)
                        try_copy_image = False
                    except:
                        wait_time += 1
                        print('error - retrying, wait_time =', wait_time)

        o.Quit()

    def copy_image(self, image_path, slide_index, img_h, img_w, horizontal_pos, vertical_pos):
        """Inserts an image into the specified slide in the PowerPoint presentation."""
        # Open the PowerPoint presentation
        presentation = Presentation(self.op_file)

        # Get the specified slide
        slide = presentation.slides[slide_index]

        # Setting Image size
        image_width = Cm(img_w)
        image_height = Cm(img_h)

        # Setting position of image in the slide
        left = Cm(horizontal_pos)
        top = Cm(vertical_pos)

        slide.shapes.add_picture(image_path, left, top, image_width, image_height)
        presentation.save(self.op_file)

    def create_folder_in_current_directory(self):
        """This function will create Images folder if does not exist"""
        folder_path = os.path.join(self.image_path)
        try:
            os.mkdir(folder_path)
            print(f"Folder '{self.image_path}' created successfully.")
        except FileExistsError:
            print(f"Folder '{self.image_path}' already exists.")

    def delete_files_in_folder(self):
        """This function will clear Images folder contents"""
        print('Delete image file -Start')
        shutil.rmtree(self.image_path)
        print('Deleted image files')

    def clear_clip_board(self):
        """This function will clear clipboard"""
        loop_count = 0
        while True:
            loop_count += 1
            try:
                ctypes.windll.user32.OpenClipboard(None)
                ctypes.windll.user32.EmptyClipboard()
                ctypes.windll.user32.CloseClipboard()
                break
            except:
                sleep(5)
                if loop_count > 5:
                    break

    def killexcel(self):
        """This function will Kill/Force Close Excel application"""
        try:
            os.system('taskkill /IM EXCEL.exe /T /F')
        except:
            pass
        sleep(5)

