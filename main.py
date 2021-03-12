# system library and defining path for module imports
import sys
sys.path.append("01_correct_script/")
sys.path.append("02_prepare_excel_script/")
sys.path.append("02_prepare_excel_script/modules")
sys.path.append("03_upload_script/")
# module import for correcting, preparing and uploading
import correct as cor
import prepare_excel as prep
import upload as up

# corrects the main excel file
cor.start()
# extract info and prepares new excel file with al the information
prep.start()
# upload the new excel file to knora
up.start()
