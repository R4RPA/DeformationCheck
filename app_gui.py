#pyinstaller --onefile --windowed --paths Lib\site-packages -i "icon.ico" app_gui.py
#pyinstaller --onefile --paths Lib\site-packages -i "icon.ico" app_gui.py <== this will keep the python terminal visible - helpful for debug and bug fix

import sys
import os
from PyQt5 import QtWidgets
from utilities.download_image2 import DownloadImage
from cf34_ui import Ui_MainWindow
from datetime import datetime

class UiWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(UiWindow, self).__init__()
        """Initiate GUI Window"""
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        """On Click and On Change Actions"""
        """Option 1 -  Extract from f59 file"""
        self.ui.browse_input_file_1.clicked.connect(self.browse_input_file_1)
        self.ui.browse_input_file_2.clicked.connect(self.browse_input_file_2)
        self.ui.browse_input_ppt_file.clicked.connect(self.browse_input_ppt_file)
        self.ui.browse_output_folder.clicked.connect(self.browse_output_folder)
        self.ui.reset_selections.clicked.connect(self.reset_selections)
        self.ui.generate_ppt.clicked.connect(self.generate_ppt)
        self.ui.exit_app.clicked.connect(self.close)

    def browse_folder(self):
        """Browse for Folder """
        return QtWidgets.QFileDialog.getExistingDirectory(self, 'Select Folder', '')

    def browse_file(self):
        """Browse for File"""
        return QtWidgets.QFileDialog.getOpenFileName(self, 'Select File', '', 'All Files (*)')[0]

    def browse_input_file_1(self):
        """Browse for input file 1"""
        file_path = self.browse_file()
        self.ui.input_file_1.setText(file_path)
        self.ui.input_file_1.setToolTip(file_path)

    def browse_input_file_2(self):
        """Browse for input file 2"""
        file_path = self.browse_file()
        self.ui.input_file_2.setText(file_path)
        self.ui.input_file_2.setToolTip(file_path)

    def browse_input_ppt_file(self):
        """Browse for template file"""
        file_path = self.browse_file()
        self.ui.input_ppt_file.setText(file_path)
        self.ui.input_ppt_file.setToolTip(file_path)
        
    def browse_output_folder(self):
        """Browse for folder to save output file"""
        folder_path = self.browse_folder()
        self.ui.output_folder.setText(folder_path)
        self.ui.output_folder.setToolTip(folder_path)
    
    def reset_selections(self):
        """Reset all extract xlife form fields to default status"""
        self.ui.input_file_1.setText("path....")
        self.ui.input_file_2.setText("path....")
        self.ui.input_ppt_file.setText("path....")
        self.ui.output_folder.setText("path....")
        self.ui.input_file_1.setToolTip("")
        self.ui.input_file_2.setToolTip("")
        self.ui.input_ppt_file.setToolTip("")
        self.ui.output_folder.setToolTip("")
        self.ui.statusbar.showMessage('')
    
    def generate_ppt(self):
        print('=== GENERATE PPT - START')
        file_paths = [self.ui.input_file_1.text(),
                   self.ui.input_file_2.text(),
                   self.ui.input_ppt_file.text(),
                   self.ui.output_folder.text()]
        
        for i, file_path in enumerate(file_paths):
            file_paths[i] = '' if file_path == "path...." else file_path

        if '' in file_paths:
            self.ui.statusbar.showMessage('MISSING Input/Output Files!')
        else:
            self.ui.statusbar.showMessage('')
            [input_file_1, input_file_2, input_ppt_file, output_folder] = file_paths
            image_path = os.path.join(os.getcwd(), "images\\")
            timestamp =  datetime.now().strftime("%Y%m%d%H%M%S")
            output_filename = f'CF34-10_LPT_Case_Deformation_check_{timestamp}.pptx'
            output_file = os.path.join(output_folder, output_filename)
            data = {"ip_file_1": input_file_1,
                    "ip_file_2": input_file_2,
                    "op_file": input_ppt_file,
                    "sv_file": output_file,
                    "image_path": image_path}

            # Initialize the class
            obj = DownloadImage(data)
            # Create directory to save the images
            obj.create_folder_in_current_directory()
            # Create ppt copy
            obj.crete_ppt()
            self.ui.statusbar.showMessage('Extracting images from Excel inputs')
            # extract images to local folder
            obj.extract_images_from_excel()
            self.ui.statusbar.showMessage('Copying the saved images and adding it to the slide')
            # copying the saved images and adding it to the slide
            obj.copy_downloaded_images_to_ppt()
            # copying the summary and judgement tables to the slide
            obj.update_pptx_tables()
            self.ui.statusbar.showMessage('Cleaning temp files created by tool')
            # deleting the files in the images1 folder
            obj.delete_files_in_folder()

            self.ui.statusbar.showMessage('PPT Generated!')
        print('=== GENERATE PPT - END')

    
def create_app():
    """Initiate PyQT Application"""
    app = QtWidgets.QApplication(sys.argv)
    win = UiWindow()
    win.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    create_app()
