from utilities.download_image2 import DownloadImage
import os

dir_path = os.getcwd()

data = {"ip_file_1": os.path.join(dir_path, r"data\CF34-10_LPT_Case_Deformation_Check3.xlsm"),
        "ip_file_2": os.path.join(dir_path, r"data\HA24K230405(JHV723EJ)径寸法データNew4.xlsm"),
        "op_file": os.path.join(dir_path, r"data\CF34-10_LPT_Case_Deformation_check_updated.pptx"),
        "image_path": os.path.join(dir_path, "images\\")}

# Initialize the class
obj = DownloadImage(data)
# Create directory to save the images
obj.create_folder_in_current_directory()
# extract images to local folder
obj.extract_images_from_excel()
# copying the saved images and adding it to the slide
obj.copy_downloaded_images_to_ppt()
# deleting the files in the images1 folder
obj.delete_files_in_folder()
