####################################################################################################
#                                                                                                  #
# Python, EMA Library handler, BBS					   						                       #
# GitHub: https://github.com/FordBBS/Handler_Excel											       #
#                                                                                                  #
####################################################################################################

#*** History ***************************************************************************************
# 2020/11/02, BBS:	- First release
# 					- Implemented 'IUser_read_excel'
# 
#***************************************************************************************************



#*** Function Group List ***************************************************************************
# - Constants & Important parameters
# - 



#*** Library Import ********************************************************************************
# Operating system
import  os

# Excel
import 	xlsxwriter
import 	pandas 			as pd
from 	openpyxl 		import load_workbook	as LoadWorkbook

#--- BBS Modules -----------------------------------------------------------------------------------
import  sys
sys.path.append(r"C:\Backup\03 SelfMade_Tools\Python\BBS_Modules")

# OS handler
import  os_handler 		as hs_os

#*** Function Group: Constants & Important parameters **********************************************
def getconst_chr_path():
	RetVal = hs_os.getconst_chr_path()
	return RetVal



#*** Function Group: Read Content ******************************************************************
def IUser_read_excel(pathFile):
	#*** Documentation *****************************************************************************
	'''Documentation,

		Read target Excel file and return in list type

		where
		[listSheetName, listAllResults]

	[str] pathFile, 	A string path of Excel file to be read

	'''

	#*** Input Validation **************************************************************************
	pathFile = str(pathFile)

	if not os.path.exists(pathFile): 	return [[], []]
	if not ".xls" in pathFile: 			return [[], []]

	#*** Initialization ****************************************************************************
	chr_path 	= getconst_chr_path()[0]
	listContent = [[], []]
	listSheet 	= []

	#*** Operations ********************************************************************************
	#--- Get list of available Worksheets ----------------------------------------------------------
	objExcel 	   = pd.ExcelFile(pathFile)
	listContent[0] = objExcel.sheet_names

	#--- Get content of each Worksheet -------------------------------------------------------------
	for thisWs in listContent[0]:
		listContent[1].append(pd.read_excel(objExcel, thisWs))
	
	#--- Release -----------------------------------------------------------------------------------
	return listContent



#*** Function Group: Write Content *****************************************************************


