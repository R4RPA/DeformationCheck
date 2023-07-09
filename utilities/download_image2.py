import win32com.client
import PIL
from PIL import ImageGrab, Image
import os
import sys
from pptx import Presentation
from pptx.util import Inches, Cm
import variables2


class DownloadImage:
    def __init__(self):
        self.current_working_dir = os.path.dirname(os.path.abspath(__file__))
        self.image_path = os.path.join(self.current_working_dir, "images\\")
        current_dir = os.getcwd()
        # self.ip_file_1 = os.path.join(current_dir,"CF34-10_LPT_Case_Deformation_Check3.xlsm")
        # self.ip_file_2 = os.path.join(current_dir,"HA24K230405(JHV723EJ)径寸法データNew4.xlsm")
        # self.op_file = os.path.join(current_dir,"CF34-10_LPT_Case_Deformation_check_updated.pptx")

        self.ip_file_1 = "C:\\Users\\1050006\\Documents\\CF34\\data\\CF34-10_LPT_Case_Deformation_Check3.xlsm"
        self.ip_file_2 = "C:\\Users\\1050006\\Documents\\CF34\\data\\HA24K230405(JHV723EJ)径寸法データNew4.xlsm"
        self.op_file = "C:\\Users\\1050006\\Documents\\CF34\\data\\CF34-10_LPT_Case_Deformation_check_updated.pptx"
        # self.ip_file = os.path.join(self.current_working_dir, "input.xlsm")
        # self.op_file = os.path.join(self.current_working_dir, "output.pptx")

    # This function extracts a graph from the input excel file and saves it into the specified PNG image path (overwrites the given PNG image)
    def save_excel_chart_as_jpg(self, filename):
        print("Processing the excel charts and saving it to the images folder in the current directory")
        # Open the excel application using win32com
        o = win32com.client.Dispatch("Excel.Application")
        # Disable alerts and visibility to the user
        o.Visible = 0
        o.DisplayAlerts = 0
        # Open workbook
        wb = o.Workbooks.Open(filename)
        for m, sheet in enumerate(o.Sheets):
            for n, shape in enumerate(sheet.Shapes):
                try:
                    shape.Copy()
                    image = ImageGrab.grabclipboard().convert('RGB')
                    new_img_path = f"{self.image_path}{sheet.name}_{n}"
                    new_img_path = new_img_path.replace(" ", "").replace(".", "") + ".jpg"
                    if os.path.exists(new_img_path):
                        new_img_path = f"{self.image_path}{sheet.name}_{n}{n}"
                        new_img_path = new_img_path.replace(" ", "").replace(".", "") + ".jpg"
                    image.save(new_img_path)
                except Exception as err:
                    # print('Failed to extract chart from sheet name : ' + str(sheet.name))
                    # print(str(err))
                    pass
        wb.Close()
        o.Quit()

    def copy_image(self, image_path, slide_index, img_h, img_w, horizontal_pos, vertical_pos):
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

    def copy_downloaded_images_to_ppt(self):
        # copying excel charts from input file and saving as jpg
        print("Process input file 1 : " + self.ip_file_1)
        self.save_excel_chart_as_jpg(self.ip_file_1)
        print("Process input file 2 : " + self.ip_file_2)
        self.save_excel_chart_as_jpg(self.ip_file_2)

        for img_file_name, val in variables2.image_filename_slide_dict.items():
            slide_idx = val["slide_index"]
            print("Inserting " + img_file_name + " into slide " + str(slide_idx) + " of file " + str(self.op_file))
            image_file = str(self.image_path + img_file_name)
            dmns = variables2.dimensions_dict[val["image_type"]]
            self.copy_image(image_file, slide_idx, dmns["image_height"], dmns["image_width"],
                            dmns["horizontal_pos"], dmns["vertical_pos"])

    def create_folder_in_current_directory(self):
        folder_path = os.path.join(self.current_working_dir, self.image_path)
        try:
            os.mkdir(folder_path)
            print(f"Folder '{self.image_path}' created successfully.")
        except FileExistsError:
            print(f"Folder '{self.image_path}' already exists.")

    def delete_files_in_folder(self):
        for filename in os.listdir(self.image_path):
            file_path = os.path.join(self.image_path, filename)
            if os.path.isfile(file_path):
                os.remove(file_path)
                print(f"Deleted file: {file_path}")


# Initialize the class
obj = DownloadImage()
# Create directory to save the images
obj.create_folder_in_current_directory()
# copying the saved images and adding it to the slide
obj.copy_downloaded_images_to_ppt()
# deleting the files in the images1 folder
obj.delete_files_in_folder()
