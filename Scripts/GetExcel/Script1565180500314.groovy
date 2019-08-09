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



//-----------------------------Getting Excel form Site-----------------------------------------------------------
//*[@id="tbl_body"]//th[10]
WebElement Tip_Out_Table = driver.findElement(By.xpath('//*[@id="tbl_body"]/tbody'))

List<WebElement> No_Of_Rows = Tip_Out_Table.findElements(By.tagName('tr'))

int Tip_Out_rows_count = No_Of_Rows.size()

WebUI.delay(100, FailureHandling.STOP_ON_FAILURE)

for (def index : (1..Tip_Out_rows_count)) {
	
	String Sno = driver.findElement(By.xpath(('//*[@id="tbl_body"]/tbody/tr[' + index) + ']/td[2]')).getText()

	String FirstName = driver.findElement(By.xpath(('//*[@id="tbl_body"]/tbody/tr[' + index) + ']/td[3]')).getText()

	String LastName = driver.findElement(By.xpath(('//*[@id="tbl_body"]/tbody/tr[' + index) + ']/td[4]')).getText()

	String EmpId = driver.findElement(By.xpath(('//*[@id="tbl_body"]/tbody/tr[' + index) + ']/td[5]')).getText()

	String Derpartment = driver.findElement(By.xpath(('//*[@id="tbl_body"]/tbody/tr[' + index) + ']/td[6]')).getText()

	String Position = driver.findElement(By.xpath(('//*[@id="tbl_body"]/tbody/tr[' + index) + ']/td[7]')).getText()

	String Shift = driver.findElement(By.xpath(('//*[@id="tbl_body"]/tbody/tr[' + index) + ']/td[8]')).getText()

	String Points = driver.findElement(By.xpath(('//*[@id="tbl_body"]/tbody/tr[' + index) + ']/td[9]')).getText()

	String HourlyRate = driver.findElement(By.xpath(('//*[@id="tbl_body"]/tbody/tr[' + index) + ']/td[10]')).getText()

	String Hours = driver.findElement(By.xpath(('//*[@id="tbl_body"]/tbody/tr[' + index) + ']/td[11]')).getText()

	String SC_Sales = driver.findElement(By.xpath(('//*[@id="tbl_body"]/tbody/tr[' + index) + ']/td[12]')).getText()

	String OT_Sales = driver.findElement(By.xpath(('//*[@id="tbl_body"]/tbody/tr[' + index) + ']/td[13]')).getText()

	String Total_Tipout = driver.findElement(By.xpath(('//*[@id="tbl_body"]/tbody/tr[' + index) + ']/td[14]')).getText()

	String SC_Tipout = driver.findElement(By.xpath(('//*[@id="tbl_body"]/tbody/tr[' + index) + ']/td[15]')).getText()

	String OT_Tipout = driver.findElement(By.xpath(('//*[@id="tbl_body"]/tbody/tr[' + index) + ']/td[16]')).getText()

	String CC_SC = driver.findElement(By.xpath(('//*[@id="tbl_body"]/tbody/tr[' + index) + ']/td[17]')).getText()

	String CC_OT = driver.findElement(By.xpath(('//*[@id="tbl_body"]/tbody/tr[' + index) + ']/td[18]')).getText()

	String excelFilePath1 = 'Data Files/Tipout.xlsx'

	workbook1 = ExcelKeywords.getWorkbook(excelFilePath1)

	ExcelKeywords.createExcelSheet(workbook1, findTestData('Company_Details').getValue(2, index1), FailureHandling.STOP_ON_FAILURE)

	sheet1 = ExcelKeywords.getExcelSheet(workbook1, findTestData('Company_Details').getValue(2, index1), FailureHandling.STOP_ON_FAILURE)

	ExcelKeywords.setValueToCellByIndex(sheet1, 0, 0, 'SNO')

	ExcelKeywords.setValueToCellByIndex(sheet1, 0, 1, 'FirstName')

	ExcelKeywords.setValueToCellByIndex(sheet1, 0, 2, 'LastName')

	ExcelKeywords.setValueToCellByIndex(sheet1, index, 0, Sno)

	ExcelKeywords.setValueToCellByIndex(sheet1, index, 1, FirstName)

	ExcelKeywords.setValueToCellByIndex(sheet1, index, 2, LastName)

	ExcelKeywords.saveWorkbook(excelFilePath1, workbook1)
	}

