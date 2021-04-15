import acm
import ael

import FRunScriptGUI

import FAddContentTab
import FPostProcessingTab
import FAdvancedSettingsTab
import FOutputSettingsTab

import FReportAPI
import FReportUtils
import FReportSheetSettingsTab
import FMacroGUI
import FLogger

logger = FLogger.FLogger( 'FAReporting' )

falseTrue = ['False','True']

class WorksheetReport(FRunScriptGUI.AelVariablesHandler):
    def getTemplates(self):
        return acm.FTradingSheetTemplate.Select('')
        
    def getPortfolios(self):
        return acm.FPhysicalPortfolio.Select('') 
    
    def __init__(self):   
        directorySelection=FRunScriptGUI.DirectorySelection()
        vars = [
                 ['template', 'Trading Sheet Template', 'FTradingSheetTemplate', self.getTemplates, None, 1, 0, 'Choose a trading sheet template', None, 1],
                 ['portfolio', 'Portfolio for report', 'FPhysicalPortfolio', self.getPortfolios, None, 1, 0,'Choose the portfolios to report', None, 1],
                 ['File Path', 'File Path', directorySelection, None, directorySelection, 1, 1, 'The file path to the directory where the report should be put. Environment variables can be specified for Windows (%VAR%) or Unix ($VAR).', None, 1]
               ]
                
        FRunScriptGUI.AelVariablesHandler.__init__(self,vars)        
           
        
ael_gui_parameters = {'windowCaption':__name__}

ael_variables = WorksheetReport()
ael_variables.LoadDefaultValues(__name__)

def ael_main(variableDictionary):
    params=FReportUtils.adjust_parameters(variableDictionary)
    report_params = FReportAPI.FWorksheetReportApiParameters()
    
    report_params.portfolios = [params['portfolio'].Name()]
    report_params.template = params['template']
    report_params.tradeRowsOnly = False
    report_params.expiredPositions = True
    report_params.zeroPositions = True
    report_params.htmlToFile = False
    report_params.htmlToScreen = False
    report_params.fileName = '_' + params['portfolio'].Name()
    report_params.fileDateFormat = '%d%m%y'
    report_params.fileDateBeginning = True
    report_params.createDirectoryWithDate = False
    report_params.overwriteIfFileExists = False
    report_params.secondaryOutput = True
    report_params.filePath = params['File Path']
    
    report_params.RunScript()

