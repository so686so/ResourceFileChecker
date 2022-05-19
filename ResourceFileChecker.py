# -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*- #
TITLE       = 'Resource File Check Program for GNET System'
VERSION     = '1.0.1'
AUTHOR      = 'So Byung Jun'
UPDATE      = '2022-5-19'
GIT_LINK    = 'https://github.com/so686so/'
# -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*- #

# Windows .exe Command : 'pyinstaller -F ResourceFileChecker.py' (need 'ResourceFileChecker.spec')
# exe 파일 패키징을 위해 코드 한곳에 통합

"""
0 : 한국어
1 : 영어
2 : 일본어
3 : 중국어
4 : 러시아어
5 : 이탈리아어
6 : 터키어
7 : 폴란드어
8 : 베트남
9 : 이스라엘
10: 프랑스어
"""

# Import Packages and Modules
# -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
# Standard Library
# -------------------------------------------------------------------
import  os
import  sys
import  inspect
import  shutil

# For Excel
# -------------------------------------------------------------------
from    openpyxl                import load_workbook

# For Xml
# -------------------------------------------------------------------
from    xml.etree.ElementTree   import  Element, SubElement
from    xml.dom                 import  minidom
import  xml.etree.ElementTree   as      ET

# For GUI ( Ver. PyQt 6.2.2 )
# -------------------------------------------------------------------
from    PyQt6                   import QtCore, QtGui, QtWidgets
from    PyQt6.QtCore            import *
from    PyQt6.QtGui             import *
from    PyQt6.QtWidgets         import *
# -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-


# Var Define
# -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
WINDOW_WIDTH        = 500
WINDOW_HEIGHT       = 500


# Const Define
# -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
__CUT_LINE__        = '=' * 54
ENCODING_FORMAT     = 'utf-8'

CUR_PATH            = os.path.dirname('./')
RESOURCE_PATH       = os.path.dirname('./resource/')

TTS_EXCEL_FILE_NAME = 'AnalysisTTS.xlsx'
XML_EXCEL_FILE_NAME = 'AnalysisLang.xlsx'

TTS_EXCEL_PATH      = os.path.join(RESOURCE_PATH, TTS_EXCEL_FILE_NAME)
XML_EXCEL_PATH      = os.path.join(RESOURCE_PATH, XML_EXCEL_FILE_NAME)

PROGRAM_LIST        = ['TTS Analysis', 'Locale XML Analysis']

PROJECT_LIST        = ['a3', 's3', 'v3', 'v4', 'v8']

XML_LANG_LIST       = ['ko' , 'en' , 'ja', 'cn', 'ru', 'it', 'tu', 'po', 'vi', 'is', 'fr']
TTS_LANG_LIST       = ['kor', 'eng', 'jp', 'cn', 'ru', 'it', 'tu', 'po', 'vi', 'is', 'fr']
GUI_SHOW_LANG_LIST  = XML_LANG_LIST

TTS                 = 0
XML                 = 1

ALIGN_CENTER        = Qt.AlignmentFlag.AlignCenter
ALIGN_LEFT          = Qt.AlignmentFlag.AlignLeft


# Define for Excel
# -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
FILE_CHECK_START    = 2
TTS_FILE_MAX        = 150
XML_FILE_MAX        = 1000


# Define for TTS
# -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
SEARCH_TTS_PATH     = os.path.join(RESOURCE_PATH, f"DefaultTTS.txt")
UNUSED_FILE_PATH    = os.path.join(RESOURCE_PATH, f"UnusedTTS.txt" )

EXTRA_TTS_LIST      = [ "touch_req.wav", ]  # For use Factory Test

""" Search Format.
    #define TTS_FILE_UPDATE_APP "update_app.wav"
"""
TTS_CHECK_STRING_1  = "#define"
TTS_CHECK_STRING_2  = "TTS_FILE"
TTS_FILE_FORMAT     = ".wav"

SUMMARY_ROW         = TTS_FILE_MAX + 3


# Define for locale XML
# -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
LOCALE_PATH         = r"source/apps/resource/language/"
ITEM_SRC            = 0
ITEM_TRAN           = 1


# Common Use Class : Signal
# -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
class LogSignal(QObject):
    signal = pyqtSignal(str)

    def sendLog(self, log) -> None:
        self.signal.emit(log)


# Common Use Class : AnalysisApp - Parent Class
# -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
class AnalysisApp:
    def __init__( self,
                  LinuxHomeDir:str,
                  ExcelFileDir:str,
                  Project:str,
                  Lang:str,
                  Signal:LogSignal
                ) -> None:
        # Set Common var
        self.BaseProject    = Project
        self.BaseLanguage   = Lang
        self.HomeDir        = LinuxHomeDir
        self.ExcelDir       = ExcelFileDir
        self.logSignal      = Signal
        self.workBook       = load_workbook(self.ExcelDir)

        # { 'a3':a3 excel sheet, 'v4':v4 excel sheet, ... }
        self.sheetNameDict  = {}

        # init
        self.initWorkBook()

    
    def showLog(self, log='', end='\n') -> None:
        self.logSignal.sendLog(log)
        print(log, end=end)


    def initWorkBook(self) -> None:
        for eachPJ in PROJECT_LIST:
            self.sheetNameDict[eachPJ] = self.workBook[eachPJ]


    # language list [ko, en, ...] -> excel sheet row index [B(66), C(67), D(68)...]
    def getExcelIndexList(self) -> list:
        return [chr(66+idx) for idx, _ in enumerate(GUI_SHOW_LANG_LIST)]


    def getSheetByProject(self, Project:str):   
        # return type - openpyxl.worksheet.worksheet.Worksheet
        return self.sheetNameDict[Project]


    def checkInitVaild(self) -> bool:
        if os.path.isdir(self.HomeDir) is False:
            self.showLog(f'! 유효한 우분투 홈 디렉토리가 아닙니다 : {self.HomeDir}')
            self.showLog(f'! 홈 디렉토리를 다시 선택해 주세요.')
            return False
        else:
            self.showLog(f'- 홈 디렉토리 : {self.HomeDir}')

        if os.paht.isfile(self.ExcelDir) is False:
            self.showLog(f'! 결과를 저장할 엑셀 파일이 유효하지 않습니다 : {self.ExcelDir}')
            return False
        else:
            self.showLog(f'- 결과 파일 : {self.ExcelDir}')

        return True

    
# class : TTS
# -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
class AnalysisTTS(AnalysisApp):
    def __init__( self, 
                  LinuxHomeDir:str, 
                  ExcelFileDir:str, 
                  Project:str, 
                  Lang:str, 
                  Signal:LogSignal
                ) -> None:
        # AnalysisApp init
        super().__init__(LinuxHomeDir, ExcelFileDir, Project, Lang, Signal)

        self.TTS_checkPath  = os.path.join(self.HomeDir, f"blackbox/{Project}/source/main/tts_out.cpp")

        self.audioDirDict   = {}
        self.curProject     = ""

        self.defaultTTSList = []
        self.curChckTTSList = []
        self.unusedFileList = []

        self.initClass()


    def initClass(self) -> bool:
        if self.set_audio_dicts() is False:
            return False

        self.load_default_search_tts_list()


    def get_audio_dir_by_project(self, Project:str) -> str:
        return f"blackbox/{Project}/source/apps/resource/audio"


    def set_audio_dicts(self) -> bool:
        for Project in PROJECT_LIST:
            audioDir = os.path.join(self.HomeDir, self.get_audio_dir_by_project(Project))

            if os.path.isdir(audioDir) is False:
                self.showLog(f" * '{audioDir}' is invalid!")
                return False

            self.audioDirDict[Project] = audioDir
            
        return True


    def load_default_search_tts_list(self) -> None:
        self.defaultTTSList = self.check_default_search_tts_list_update()


    def check_default_search_tts_list_update(self) -> bool:
        if os.path.isfile(self.TTS_checkPath) is False:
            self.showLog(f" * '{self.TTS_checkPath}' is Invalid! Program Quit!")
            sys.exit(-1)

        FILE_NAME_SPLIT       = 1 # Const
        origin_TTS_list_count = 0
        update_TTS_list_count = 0
        origin_list           = []
        update_list           = []

        # Load realTime tts_out.cpp wav files
        with open(self.TTS_checkPath, 'r', encoding=ENCODING_FORMAT) as rf:
            for eachLine in rf:
                eachLine = eachLine.strip('\n')

                # Check Line Format, like '#define TTS_FILE_UPDATE_APP "update_app.wav"'
                if  TTS_CHECK_STRING_1 in eachLine and \
                    TTS_CHECK_STRING_2 in eachLine :
                    TTS_FileName = eachLine.split('\"')[FILE_NAME_SPLIT]

                    # Check ext .wav
                    if  TTS_FILE_FORMAT in TTS_FileName:
                        update_list.append(TTS_FileName)

        update_list.extend(EXTRA_TTS_LIST)
        update_TTS_list_count = len( update_list )

        # Check Before tts search file list
        with open(SEARCH_TTS_PATH, 'r', encoding=ENCODING_FORMAT) as rf:
            for eachLine in rf:
                origin_list.append(eachLine.strip('\n'))
                origin_TTS_list_count += 1

        self.showLog()
        self.showLog(__CUT_LINE__)
        self.showLog('* Check Default Search TTS file is updated.')
        self.showLog(__CUT_LINE__)
        self.showLog(f'- Origin TTS File Count : {origin_TTS_list_count}')
        self.showLog(f'- Update TTS File Count : {update_TTS_list_count}')
        self.showLog(__CUT_LINE__)

        originSet = set(origin_list)
        updateSet = set(update_list)

        if originSet != updateSet:
            self.showLog(f'* Update Search TTS File List - {originSet - updateSet}\n')

            with open(SEARCH_TTS_PATH, 'w', encoding=ENCODING_FORMAT) as wf:
                for line in update_list:
                    wf.write(f'{line}\n')
            return update_list

        else:
            self.showLog('* File Matched\n')
            return origin_list


    def set_currentProject(self, Project:str) -> None:
        self.showLog(f'* Current Project : {Project.capitalize()}')
        self.curProject = Project


    def load_TTS_list_by_language(self, Lang:str) -> bool:
        curAudioDir = self.audioDirDict[self.curProject]
        curLangDir  = os.path.join(curAudioDir, Lang)

        self.curChckTTSList.clear()

        if os.path.isdir(curLangDir) is False:
            self.showLog(f'* 유효하지 않은 폴더입니다 : {curLangDir}')
            return False

        for eachFile in os.listdir(curLangDir):
            self.curChckTTSList.append(eachFile)

        return True


    def run(self):
        self.showLog()
        self.showLog(__CUT_LINE__)
        self.showLog('* Run Program - TTS')
        self.showLog(__CUT_LINE__)
        self.showLog(f'* Base Project  : {self.BaseProject}')
        self.showLog('* Base Language : Ko ( Fixed when TTS )')
        self.showLog(__CUT_LINE__)
        self.showLog()

        for Project in PROJECT_LIST:
            self.set_currentProject(Project)

            # Const Define
            DEFAULT_COL_WORD    = 'A'
            TITLE               = 1
            EXIST_MARK          = 'O'

            selectSheet         = self.getSheetByProject(Project)
            returnDict          = {}
            ExcelSheetRowIdx    = self.getExcelIndexList()

            selectSheet['A1']   = 'DefaultFile'

            self.unusedFileList.append(f'[ {Project} ] - Files that are not used compared to DefualtTTS')
            self.unusedFileList.append('---------------------------------------------------------------')

            overwrittenCheckSet = set()

            # Set Default wav File name 'A' Col
            for idx, eachDefault in enumerate(self.defaultTTSList):
                Cell = f'{DEFAULT_COL_WORD}{idx + FILE_CHECK_START}'
                selectSheet[Cell] = str(eachDefault)

            for LangIdx, Language in enumerate(TTS_LANG_LIST):
                isReadFiles = self.load_TTS_list_by_language(Language)

                Cell = f"{ExcelSheetRowIdx[LangIdx]}{TITLE}"
                selectSheet[Cell] = Language

                fileCount = 0

                self.unusedFileList.append(f'<{Language}>')

                if isReadFiles is True:
                    overwrittenCheckSet = set(self.curChckTTSList.copy())

                    for searchIdx, eachDefault in enumerate(self.defaultTTSList):
                        Cell = f"{ExcelSheetRowIdx[LangIdx]}{searchIdx + FILE_CHECK_START}"

                        if eachDefault in self.curChckTTSList:
                            selectSheet[Cell] = EXIST_MARK
                            overwrittenCheckSet.discard(eachDefault)
                            fileCount += 1

                        else:
                            selectSheet[Cell] = ''

                self.showLog(f' # {Language:3} - {fileCount}개')

                for eachUnused in overwrittenCheckSet:
                    self.unusedFileList.append(f' - {eachUnused}')
                self.unusedFileList.append('')

                # TTS_FILE_MAX

                ClearStartLine = len( self.defaultTTSList ) + FILE_CHECK_START

                for row in selectSheet[f'{DEFAULT_COL_WORD}{ClearStartLine}:{ExcelSheetRowIdx[-1]}{TTS_FILE_MAX}']:
                    for cell in row:
                        cell.value = None

                for idx, _ in enumerate(TTS_LANG_LIST):
                    Word = ExcelSheetRowIdx[idx]
                    Cell = f"{Word}{SUMMARY_ROW}"
                    selectSheet[Cell] = f'=COUNTIF({Word}{FILE_CHECK_START}:{Word}{ClearStartLine}, "{EXIST_MARK}")'

            self.showLog()

        with open(UNUSED_FILE_PATH, 'w', encoding=ENCODING_FORMAT) as wf:
            for line in self.unusedFileList:
                wf.write(f'{line}\n')

        self.showLog(__CUT_LINE__)
        self.showLog(f'Excel File Save... : {self.ExcelDir}')
        self.workBook.save(self.ExcelDir)
        self.workBook.close()

        self.showLog('[!] TTS Anlysis Done!')
        self.showLog(__CUT_LINE__)
        self.showLog()


# class : XML
# -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
class AnalysisXML(AnalysisApp):
    def __init__( self, 
                  LinuxHomeDir:str, 
                  ExcelFileDir:str, 
                  Project:str, 
                  Lang:str, 
                  Signal:LogSignal
                ) -> None:
        # AnalysisApp init
        super().__init__(LinuxHomeDir, ExcelFileDir, Project, Lang, Signal)

        self.cur_project        = ""

        """ [ parse_dict Format ]
            {
                Project: {
                    Language : {
                        context id : [
                            ( src, tran ), ...
                        ]
                        ...
                    }
                    ...
                }
                ...
            }
        """
        self.parse_dict         = {}
        self.base_comp_dict     = {}
        self.base_comp_count    = 0
        self.base_comp_val_list = []

        self.initSetPath()


    def getLanguageDirPath_by_Project(self, Project:str) -> str:
        return os.path.join(self.HomeDir, f"blackbox/{Project}/{LOCALE_PATH}")

    
    def getLocaleFilePath_by_Project(self, Project:str, Lang:str) -> str:
        return os.path.join(self.getLanguageDirPath_by_Project(Project), f"locale_{Lang}.xml")


    def checkExistLocaleFile_by_Project(self, Project:str) -> None:
        langDir = self.getLanguageDirPath_by_Project(Project)
        logString = ""

        self.parse_dict[Project] = {}

        if os.path.isdir(langDir) is False:
            self.showLog(f'- {Project.capitalize()} : X')
            return

        logString += f"- {Project.capitalize()} :"
        for eachLang in XML_LANG_LIST:
            self.parse_dict[Project][eachLang] = {}
            if os.path.isfile(self.getLocaleFilePath_by_Project(Project, eachLang)) is True:
                logString += f" {eachLang}[O]"
            else:
                logString += f" {eachLang}[X]"
        self.showLog(logString)


    def initSetPath(self) -> None:
        self.showLog(f'\n* Initialize')
        self.showLog(__CUT_LINE__)
        for Project in PROJECT_LIST:
            self.checkExistLocaleFile_by_Project(Project)


    def ParseXml(self, Project:str, Lang:str) -> dict:
        try:
            tree    = ET.parse(self.getLocaleFilePath_by_Project(Project, Lang))
        except Exception:
            return

        root        = tree.getroot()
        returnDict  = {}

        for context in root.findall('context'):
            contextID   = context.attrib.get('id')
            items       = context.findall('item')

            itemList    = []
            for item in items:
                itemList.append((item.attrib.get('src'), item.attrib.get('tran')))

            returnDict[contextID] = itemList

        return returnDict


    def setBaseDict(self):
        self.showLog()
        self.showLog(__CUT_LINE__)
        self.showLog(f'* SetBaseDict Start')
        self.base_comp_dict = self.parse_dict[self.BaseProject][self.BaseLanguage]
        self.calcBaseCompareDictCount()
        self.setBaseValueList()

        self.showLog(f'* SetBaseDict Done : {self.BaseProject} - {self.BaseLanguage}')
        self.showLog(f'- TotalCount : {self.base_comp_count}개')
        self.showLog(__CUT_LINE__)


    def calcBaseCompareDictCount(self):
        count = 0
        for _, v in self.base_comp_dict.items():
            count += len(v)
        self.base_comp_count = count


    def setBaseValueList(self):
        totalList = []
        for _, v in self.base_comp_dict.items():
            totalList.extend(v)
        self.base_comp_val_list = totalList


    def getTotalItemList_by_Project_Language(self, Project:str, Language:str) -> list:
        totalList = []
        if not self.parse_dict[Project][Language]:
            return totalList

        for _, v in self.parse_dict[Project][Language].items():
            totalList.extend(v)

        return totalList


    def checkMatchedPercentage_by_Language(self, Project:str, Language:str) -> float:
        calcPercent = 0.0
        calcCount   = 0

        compList    = [ v[ITEM_SRC] for v in self.base_comp_val_list ]
        itemList    = [ v[ITEM_SRC] for v in self.getTotalItemList_by_Project_Language(Project, Language) ]

        for item in itemList:
            if item in compList:
                calcCount += 1

        calcPercent = (calcCount/ self.base_comp_count) * 100
        return round(calcPercent, 2)


    def run(self):
        self.showLog()
        self.showLog(__CUT_LINE__)
        self.showLog('* Run Program - XML')
        self.showLog(__CUT_LINE__)
        self.showLog(f'* Base Project  : {self.BaseProject}')
        self.showLog(f'* Base Language : {self.BaseLanguage}')
        self.showLog(__CUT_LINE__)
        self.showLog()

        self.showLog()
        self.showLog(__CUT_LINE__)
        self.showLog('ParseXML Start')
        self.showLog(__CUT_LINE__)
        for Project in PROJECT_LIST:
            self.showLog(f'- Parsing {Project.capitalize()}...')

            for Language in XML_LANG_LIST:
                self.parse_dict[Project][Language] = self.ParseXml(Project, Language)
        self.showLog(__CUT_LINE__)

        self.setBaseDict()

        self.showLog()
        self.showLog(__CUT_LINE__)
        self.showLog('CheckXML Start')
        self.showLog(__CUT_LINE__)

        for Project in PROJECT_LIST:
            self.showLog(f'- Check {Project.capitalize()}...')
            for Language in XML_LANG_LIST:
                self.showLog(f'  > {Language} : {self.checkMatchedPercentage_by_Language(Project, Language)}%')
            self.showLog('\n')
        self.showLog(__CUT_LINE__)

        self.saveExcel()


    def saveExcel(self):
        self.showLog()
        self.showLog(__CUT_LINE__)
        self.showLog(f'* Save Result Excel... : {self.ExcelDir}')
        self.showLog(__CUT_LINE__)

        wordList = self.getExcelIndexList()

        for Project in PROJECT_LIST:
            self.showLog(f'- Save {Project.capitalize()}...')
            selectSheet = self.getSheetByProject(Project)

            for row in selectSheet[f'A1:{wordList[-1]}{XML_FILE_MAX}']:
                for cell in row:
                    cell.value = None

            selectSheet['A1'] = f'Source Word[{self.BaseLanguage}]'

            compList = [ v[ITEM_SRC] for v in self.base_comp_val_list ]

            for idx, eachItem in enumerate(self.base_comp_val_list):
                Cell = f'A{idx + FILE_CHECK_START}'
                selectSheet[Cell] = eachItem[ITEM_SRC]

            for LangIdx, Language in enumerate(XML_LANG_LIST):
                rowWord = wordList[LangIdx]
                selectSheet[f'{rowWord}1'] = f'=CONCATENATE("{Language}(",ROUND((COUNTA({rowWord}{FILE_CHECK_START}:{rowWord}{XML_FILE_MAX})/COUNTA(A2:A{XML_FILE_MAX}))*100,2)," %)")'

                ItemList = [ v[ITEM_SRC]  for v in self.getTotalItemList_by_Project_Language(Project, Language) ]
                TranList = [ v[ITEM_TRAN] for v in self.getTotalItemList_by_Project_Language(Project, Language) ]

                for compIdx, eachCompItem in enumerate(compList):
                    for iIdx, eachItem in enumerate(ItemList):
                        if eachItem == eachCompItem:
                            Cell = f'{rowWord}{compIdx + FILE_CHECK_START}'
                            selectSheet[Cell] = TranList[iIdx]

        self.showLog(__CUT_LINE__)
        self.showLog(f'* All Project Excel Write Done')
        self.showLog(__CUT_LINE__)
        self.showLog('- Try Save...')

        self.workBook.save(self.ExcelDir)
        self.workBook.close()

        self.showLog('- Save Done')
        self.showLog(__CUT_LINE__)
        self.showLog('[!] XML Anlysis Done!')
        self.showLog(__CUT_LINE__)
        self.showLog()


# Ui_MainWindow Class (MainWindow)
# -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
class Ui_MainWindow(object):
    def setupUi(self, MainWindow:QtWidgets.QMainWindow):
        # 전체
        MainWindow.setObjectName('MainWindow')
        MainWindow.setWindowTitle('Resource File Checker')

        MainWindow.resize(WINDOW_WIDTH, WINDOW_HEIGHT)
        MainWindow.setFixedSize(WINDOW_WIDTH, WINDOW_HEIGHT)

        self.centralWidget          = QtWidgets.QWidget(MainWindow)

        # 상단
        self.selectProgramGroupBox  = QtWidgets.QGroupBox('Select', MainWindow)

        self.selectProgramComboBox  = QtWidgets.QComboBox(MainWindow)
        self.selectProjectComboBox  = QtWidgets.QComboBox(MainWindow)
        self.selectLanguageComboBox = QtWidgets.QComboBox(MainWindow)
        self.selectProgramLabel     = QtWidgets.QLabel('실행할 프로그램')
        self.selectInformationLabel = QtWidgets.QLabel('프로그램 실행시 기준값')

        self.top_V_Layout           = QtWidgets.QVBoxLayout()
        self.top_upper_H_Layout     = QtWidgets.QHBoxLayout()
        self.top_under_H_Layout     = QtWidgets.QHBoxLayout()

        self.top_upper_H_Layout.addWidget(self.selectProgramLabel, 1)
        self.top_upper_H_Layout.addWidget(self.selectProgramComboBox, 2)
        self.top_under_H_Layout.addWidget(self.selectInformationLabel, 3)
        self.top_under_H_Layout.addWidget(self.selectProjectComboBox, 1)
        self.top_under_H_Layout.addWidget(self.selectLanguageComboBox, 1)

        self.top_V_Layout.addLayout(self.top_upper_H_Layout)
        self.top_V_Layout.addLayout(self.top_under_H_Layout)

        self.selectProgramGroupBox.setLayout(self.top_V_Layout)

        # 상단 세팅
        self.selectProgramComboBox.addItems(PROGRAM_LIST)
        self.selectProjectComboBox.addItems(PROJECT_LIST)
        self.selectLanguageComboBox.addItems(GUI_SHOW_LANG_LIST)

        self.selectProjectComboBox.setCurrentIndex(3)


        # 중단
        self.selectPathGroupBox     = QtWidgets.QGroupBox('Directory', MainWindow)

        self.linuxHomeDirLabel      = QtWidgets.QLabel('우분투 홈 디렉토리')
        self.linuxHomeDirLineEdit   = QtWidgets.QLineEdit('C:/')
        self.linuxHomeDirButton     = QtWidgets.QPushButton('Edit')

        self.resultDirLabel         = QtWidgets.QLabel('엑셀 결과 저장 폴더')
        self.resultDirLineEdit      = QtWidgets.QLineEdit(CUR_PATH)
        self.resultDirButton        = QtWidgets.QPushButton('Edit')

        self.mid_V_Layout           = QtWidgets.QVBoxLayout()
        self.mid_upper_H_Layout     = QtWidgets.QHBoxLayout()
        self.mid_under_H_Layout     = QtWidgets.QHBoxLayout()

        self.mid_upper_H_Layout.addWidget(self.linuxHomeDirLabel, 2)
        self.mid_upper_H_Layout.addWidget(self.linuxHomeDirLineEdit, 3)
        self.mid_upper_H_Layout.addWidget(self.linuxHomeDirButton, 1)

        self.mid_under_H_Layout.addWidget(self.resultDirLabel, 2)
        self.mid_under_H_Layout.addWidget(self.resultDirLineEdit, 3)
        self.mid_under_H_Layout.addWidget(self.resultDirButton, 1)

        self.mid_V_Layout.addLayout(self.mid_upper_H_Layout)
        self.mid_V_Layout.addLayout(self.mid_under_H_Layout)

        self.selectPathGroupBox.setLayout(self.mid_V_Layout)

        # 중단 세팅
        self.linuxHomeDirLineEdit.setReadOnly(True)
        self.resultDirLineEdit.setReadOnly(True)


        # 하단
        self.bot_H_Layout           = QtWidgets.QHBoxLayout()

        self.LogTextEdit            = QtWidgets.QTextEdit()
        self.ClearButton            = QtWidgets.QPushButton('Log Clear')
        self.RunButton              = QtWidgets.QPushButton('Run')

        self.bot_H_Layout.addWidget(self.ClearButton, 1)
        self.bot_H_Layout.addWidget(self.RunButton,   1)

        self.mainLayout = QtWidgets.QVBoxLayout()
        self.mainLayout.addWidget(self.selectProgramGroupBox, 1)
        self.mainLayout.addWidget(self.selectPathGroupBox, 1)
        self.mainLayout.addWidget(self.LogTextEdit, 5)
        self.mainLayout.addLayout(self.bot_H_Layout, 1)

        self.centralWidget.setLayout(self.mainLayout)
        MainWindow.setCentralWidget(self.centralWidget)


# AnalysisAppUI Class
# -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
class AnalysisAppUI(QMainWindow):
    def __init__(self, QApp=None):
        super().__init__()
        self.app                = QApp

        self.PreRememberPath    = CUR_PATH
        self.copiedResultExcel  = r""
        self.LinuxHomeDir       = r""
        self.ResultDir          = r""

        self.logSignal          = LogSignal()
        self.ui                 = Ui_MainWindow()
        self.ui.setupUi(self)

        self.initialize()


    @pyqtSlot(str)
    def TRACE(self, log='', align=ALIGN_LEFT):
        if log == __CUT_LINE__:
            self.ui.LogTextEdit.append(log)
            self.ui.LogTextEdit.setAlignment(ALIGN_CENTER)
        else:
            self.ui.LogTextEdit.append(log)
            self.ui.LogTextEdit.setAlignment(align)
            

    def setDefaultLog(self):
        self.TRACE(__CUT_LINE__)
        self.ui.LogTextEdit.setFontPointSize(13)
        self.TRACE(TITLE, ALIGN_CENTER)
        self.ui.LogTextEdit.setFontPointSize(9)
        self.TRACE(f'ver. {VERSION} / {UPDATE}', ALIGN_CENTER)
        self.TRACE()
        self.TRACE(AUTHOR, ALIGN_CENTER)
        self.TRACE(f'({GIT_LINK})', ALIGN_CENTER)
        self.TRACE(__CUT_LINE__)
        self.TRACE()
        self.TRACE('  [ How To Run ]')
        self.TRACE(__CUT_LINE__)
        self.TRACE('  1. 우분투 홈 디렉토리를 자신의 우분투 홈 디렉토리로 설정')
        self.TRACE('  2. 엑셀 결과 저장 폴더를 결과 엑셀 파일을 받고싶은 디렉토리로 설정')
        self.TRACE('  3. 두 디렉토리가 모두 유효한 디렉토리라면 하단의 Run 버튼 활성화')
        self.TRACE(__CUT_LINE__)
        self.TRACE()


    def initialize(self):
        self.ui.linuxHomeDirButton.clicked.connect(self.selectHomeDir)
        self.ui.resultDirButton.clicked.connect(self.selectResultDir)
        self.ui.RunButton.clicked.connect(self.runProgram)
        self.ui.ClearButton.clicked.connect(self.clearLogEdit)
        self.ui.selectProgramComboBox.currentIndexChanged.connect(self.checkComboboxProgram)
        self.ui.selectProjectComboBox.currentIndexChanged.connect(self.checkComboboxProject)
        self.ui.selectLanguageComboBox.currentIndexChanged.connect(self.checkComboboxLanguage)
        self.logSignal.signal.connect(self.TRACE)

        self.setDefaultLog()
        self.checkComboboxProgram()
        self.ui.RunButton.setDisabled(True)


    def checkActiveRunButton(self):
        if  os.path.isdir(os.path.join(self.LinuxHomeDir, 'blackbox')) is True and \
            os.path.isdir(self.ResultDir) is True:
            self.TRACE(__CUT_LINE__)
            self.TRACE('[!] 디렉토리 세팅 완료 : 이제 프로그램을 실행할 수 있습니다.')
            self.TRACE(__CUT_LINE__)
            self.ui.RunButton.setEnabled(True)
        else:
            self.ui.RunButton.setDisabled(True)


    def checkComboboxProgram(self):
        curSelect = self.ui.selectProgramComboBox.currentText()

        self.TRACE(f"- Program Select : {curSelect}")
        if curSelect == 'TTS Analysis':
            self.ui.selectLanguageComboBox.setCurrentIndex(0)
            self.ui.selectLanguageComboBox.setDisabled(True)
        else:
            self.ui.selectLanguageComboBox.setEnabled(True)


    def checkComboboxProject(self):
        curProject = self.ui.selectProjectComboBox.currentText()
        self.TRACE(f"- Project Changed : {curProject}")


    def checkComboboxLanguage(self):
        curLang = self.ui.selectLanguageComboBox.currentText()
        self.TRACE(f"- Language Changed : {curLang}")


    def selectHomeDir(self):
        sender      = self.sender()
        prePath     = self.ui.linuxHomeDirLineEdit.text()
        targetDir   = QFileDialog.getExistingDirectory(self, 'Select Path', self.PreRememberPath)

        if len(targetDir) == 0:
            targetDir = prePath
        else:
            if os.path.isdir(os.path.join(targetDir, 'blackbox')) is True:
                self.TRACE(__CUT_LINE__)
                self.TRACE(f'* 홈 디렉토리 세팅 완료 : {targetDir}')
                self.TRACE(__CUT_LINE__)
            else:
                self.TRACE(__CUT_LINE__)
                self.TRACE(f'* 유효하지 않은 홈 디렉토리입니다. : {targetDir}')
                self.TRACE('* 우분투 홈 디렉토리를 다시 선택해 주세요.')
                self.TRACE(__CUT_LINE__)
                targetDir = prePath

        self.PreRememberPath = os.path.dirname(targetDir)
        self.ui.linuxHomeDirLineEdit.setText(targetDir)
        self.LinuxHomeDir = targetDir

        self.checkActiveRunButton()


    def selectResultDir(self):
        sender      = self.sender()
        prePath     = self.ui.resultDirLineEdit.text()
        targetDir   = QFileDialog.getExistingDirectory(self, 'Select Path', CUR_PATH)

        if len(targetDir) == 0:
            targetDir = prePath
        else:
            self.TRACE(__CUT_LINE__)
            self.TRACE(f'* 결과 디렉토리 세팅 완료 : {targetDir}')
            self.TRACE(__CUT_LINE__)

        self.ui.resultDirLineEdit.setText(targetDir)
        self.ResultDir = targetDir

        self.checkActiveRunButton()


    def clearLogEdit(self):
        self.ui.LogTextEdit.clear()
        self.setDefaultLog()


    def runProgram(self):
        ProgramName = self.ui.selectProgramComboBox.currentText()
        ProjectName = self.ui.selectProjectComboBox.currentText()
        LangName    = self.ui.selectLanguageComboBox.currentText()

        self.ui.RunButton.setDisabled(True)

        self.TRACE('\n>>>>>>>>>>>>>>>>> [RUN]')
        self.TRACE(__CUT_LINE__)
        self.TRACE(f'Program : {ProgramName}')
        self.TRACE(f'Project : {ProjectName}')
        self.TRACE(f'Lang    : {LangName}')
        self.TRACE(__CUT_LINE__)

        if ProgramName == 'TTS Analysis':
            self.runTTS(ProjectName, LangName)

        elif ProgramName == 'Locale XML Analysis':
            self.runXML(ProjectName, LangName)

        os.startfile(self.ResultDir)

        self.ui.RunButton.setEnabled(True)


    def runTTS(self, Project:str, Lang:str):
        for idx, eachLang in enumerate(GUI_SHOW_LANG_LIST):
            if Lang == eachLang:
                Lang = TTS_LANG_LIST[idx]
                self.BaseLanguage = Lang


        self.copiedResultExcel = os.path.join(self.ResultDir, 'AnalysisTTS.xlsx')
        shutil.copyfile(TTS_EXCEL_PATH, self.copiedResultExcel)

        if os.path.isfile(self.copiedResultExcel) is True:
            self.TRACE(__CUT_LINE__)
            self.TRACE('* Setting : Copy Excel File Done')
            self.TRACE(__CUT_LINE__)
        else:
            self.TRACE(__CUT_LINE__)
            self.TRACE('* Setting : Copy Excel File Fail')
            self.TRACE(__CUT_LINE__)
            return

        appTTS = AnalysisTTS(self.LinuxHomeDir, self.copiedResultExcel, Project, Lang, self.logSignal)
        appTTS.run()


    def runXML(self, Project:str, Lang:str):
        for idx, eachLang in enumerate(GUI_SHOW_LANG_LIST):
            if Lang == eachLang:
                Lang = XML_LANG_LIST[idx]
                self.BaseLanguage = Lang

        self.copiedResultExcel = os.path.join(self.ResultDir, 'AnalysisLang.xlsx')
        shutil.copyfile(XML_EXCEL_PATH, self.copiedResultExcel)

        if os.path.isfile(self.copiedResultExcel) is True:
            self.TRACE(__CUT_LINE__)
            self.TRACE('* Setting : Copy Excel File Done')
            self.TRACE(__CUT_LINE__)
        else:
            self.TRACE(__CUT_LINE__)
            self.TRACE('* Setting : Copy Excel File Fail')
            self.TRACE(__CUT_LINE__)
            return

        appXML = AnalysisXML(self.LinuxHomeDir, self.copiedResultExcel, Project, Lang, self.logSignal)
        appXML.run()


    def run(self):
        self.show()
        self.app.exec()


if __name__ == "__main__":
    App             = QApplication(sys.argv)
    AnalysisApp     = AnalysisAppUI(App)
    AnalysisApp.run()