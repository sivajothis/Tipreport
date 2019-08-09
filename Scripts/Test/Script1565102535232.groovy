import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import internal.GlobalVariable as GlobalVariable
import com.kms.katalon.keyword.excel.ExcelKeywords as ExcelKeywords
import com.kms.katalon.core.util.KeywordUtil


WebUI.openBrowser('')

String[] sites = ['http://jetinfo.us', 'http://kkdi.jetinfo.us']

WebUI.navigateToUrl('http://jetinfo.us/')

String company = findTestData('Company_Details').getValue(1, 1)

String excelFilePath1 = 'Data Files/Tipout.xlsx'

workbook1 = ExcelKeywords.getWorkbook(excelFilePath1)

ExcelKeywords.createExcelSheet(workbook1, company)

sheet1 = ExcelKeywords.getExcelSheet(workbook1, company)

cellA1 = ExcelKeywords.getCellByIndex(sheet1, 1, 0)

cellA2 = ExcelKeywords.getCellByIndex(sheet1, 2, 0)

def liveValues = ExcelKeywords.getCellValue(cellA1)

def TestValues = ExcelKeywords.getCellValue(cellA2)

println('First row' + liveValues)

println('Secound row' + TestValues)

ExcelKeywords.compareTwoCells(liveValues, TestValues)

if(liveValues != TestValues) {
	
    KeywordUtil.logInfo(liveValues + " and " + TestValues + " do not match.")
}

