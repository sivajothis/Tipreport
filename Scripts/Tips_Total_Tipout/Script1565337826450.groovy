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
import com.kms.katalon.core.webui.driver.DriverFactory as DriverFactory
import org.openqa.selenium.WebDriver as WebDriver
import org.openqa.selenium.Keys as Keys
import org.openqa.selenium.WebElement as WebElement
import org.openqa.selenium.By as By
import com.kms.katalon.core.testobject.ConditionType as ConditionType
import com.kms.katalon.core.util.KeywordUtil as KeywordUtil

WebUI.openBrowser('')

String[] sites = ['http://jetinfo.us', 'http://kkdi.jetinfo.us']

sitecount = sites.length

for (def index : (0..sitecount - 1)) {
	
    WebUI.navigateToUrl(sites[index])

    WebDriver driver = DriverFactory.getWebDriver()

    String excelFilePath = 'Data Files/login.xlsx'

    sheet01 = ExcelKeywords.getExcelSheet(excelFilePath, 0)

    Rowcount = sheet01.physicalNumberOfRows
	
	WebUI.callTestCase(findTestCase('Tips_login'), [:], FailureHandling.STOP_ON_FAILURE)

    for (def index1 : (1..Rowcount - 1)) {			
		
        WebUI.mouseOver(findTestObject('Object Repository/Omnivore'))

        WebUI.click(findTestObject('Object Repository/Confirm_Edit'))

        String company = findTestData('Company_Details').getValue(1, index1)

        TestObject selectBox = new TestObject('SelectBox').addProperty('css', ConditionType.EQUALS, 'input.select-dropdown')

        WebUI.click(selectBox)

        TestObject dropdownValue = new TestObject('DropDownValue').addProperty('xpath', ConditionType.EQUALS, ('//ul[contains(@class, \'dropdown-content select-dropdown\')]/li/span[contains(text(),\'' + 
            company) + '\')]')

        WebUI.waitForElementVisible(dropdownValue, 30)

        WebUI.click(dropdownValue)

        String date = findTestData('Company_Details').getValue(2, index1)

        WebUI.setText(findTestObject('Calendar'), date)

        WebUI.sendKeys(findTestObject('Object Repository/input_Date_dp1'), Keys.chord(Keys.ENTER))

        WebUI.delay(5)

        WebUI.click(findTestObject('Tip_Out'))

        WebUI.click(findTestObject('Pop_Up_Yes'))

        WebUI.waitForElementVisible(findTestObject('TIP_Wait'), 1000)

        WebUI.delay(40)

        String Hours = driver.findElement(By.xpath('//*[@id="tbl_body"]/thead/tr[3]/th[1]')).getText()

        String SC = driver.findElement(By.xpath('//*[@id="tbl_body"]/thead/tr[3]/th[2]')).getText()

        String OT = driver.findElement(By.xpath('//*[@id="tbl_body"]/thead/tr[3]/th[3]')).getText()

        String Total_Tip = driver.findElement(By.xpath('//*[@id="tbl_body"]/thead/tr[3]/th[4]')).getText()

        String SC_Tip = driver.findElement(By.xpath('//*[@id="tbl_body"]/thead/tr[3]/th[5]')).getText()

        String OT_Tip = driver.findElement(By.xpath('//*[@id="tbl_body"]/thead/tr[3]/th[6]')).getText()

        String CC_SC = driver.findElement(By.xpath('//*[@id="tbl_body"]/thead/tr[3]/th[7]')).getText()

        String CC_OT = driver.findElement(By.xpath('//*[@id="tbl_body"]/thead/tr[3]/th[8]')).getText()

        String excelFilePath1 = 'Data Files/Tipout.xlsx'

        workbook1 = ExcelKeywords.getWorkbook(excelFilePath1)

        ExcelKeywords.createExcelSheet(workbook1, company)

        sheet1 = ExcelKeywords.getExcelSheet(workbook1, company)

        if (index == 0) {
            ExcelKeywords.setValueToCellByIndex(sheet1, 0, 0, 'Hours')

            ExcelKeywords.setValueToCellByIndex(sheet1, 0, 1, 'SC')

            ExcelKeywords.setValueToCellByIndex(sheet1, 0, 2, 'OT')

            ExcelKeywords.setValueToCellByIndex(sheet1, 0, 3, 'Total_Tip')

            ExcelKeywords.setValueToCellByIndex(sheet1, 0, 4, 'SC_Tip')

            ExcelKeywords.setValueToCellByIndex(sheet1, 0, 5, 'OT_Tip')

            ExcelKeywords.setValueToCellByIndex(sheet1, 0, 6, 'CC_SC')

            ExcelKeywords.setValueToCellByIndex(sheet1, 0, 7, 'CC_OT')

            ExcelKeywords.setValueToCellByIndex(sheet1, 1, 0, Hours)

            ExcelKeywords.setValueToCellByIndex(sheet1, 1, 1, SC)

            ExcelKeywords.setValueToCellByIndex(sheet1, 1, 2, OT)

            ExcelKeywords.setValueToCellByIndex(sheet1, 1, 3, Total_Tip)

            ExcelKeywords.setValueToCellByIndex(sheet1, 1, 4, SC_Tip)

            ExcelKeywords.setValueToCellByIndex(sheet1, 1, 5, OT_Tip)

            ExcelKeywords.setValueToCellByIndex(sheet1, 1, 6, CC_SC)

            ExcelKeywords.setValueToCellByIndex(sheet1, 1, 7, CC_OT)

            ExcelKeywords.saveWorkbook(excelFilePath1, workbook1)
        } else {
            ExcelKeywords.setValueToCellByIndex(sheet1, 2, 0, Hours)

            ExcelKeywords.setValueToCellByIndex(sheet1, 2, 1, SC)

            ExcelKeywords.setValueToCellByIndex(sheet1, 2, 2, OT)

            ExcelKeywords.setValueToCellByIndex(sheet1, 2, 3, Total_Tip)

            ExcelKeywords.setValueToCellByIndex(sheet1, 2, 4, SC_Tip)

            ExcelKeywords.setValueToCellByIndex(sheet1, 2, 5, OT_Tip)

            ExcelKeywords.setValueToCellByIndex(sheet1, 2, 6, CC_SC)

            ExcelKeywords.setValueToCellByIndex(sheet1, 2, 7, CC_OT)

            ExcelKeywords.saveWorkbook(excelFilePath1, workbook1)

            for (def index3 : (0..7)) {
				
                cellA0 = ExcelKeywords.getCellByIndex(sheet1, 0, index3)

                cellA1 = ExcelKeywords.getCellByIndex(sheet1, 1, index3)

                cellA2 = ExcelKeywords.getCellByIndex(sheet1, 2, index3)

                def liveValues = ExcelKeywords.getCellValue(cellA1)

                def TestValues = ExcelKeywords.getCellValue(cellA2)

                def Iteams = ExcelKeywords.getCellValue(cellA0)
			
                if (liveValues != TestValues) {
					
                    KeywordUtil.markFailed(((((((((Iteams + ' ') + company) + ' ') + date) + ' live site value is ') + liveValues) + 
                        ' and Test Site ') + TestValues) + ' do not match.')
                } else {
                    KeywordUtil.markPassed(((((((((Iteams + ' ') + company) + ' ') + date) + ' live site value is ') + liveValues) + 
                        ' and Test Site ') + TestValues) + ' do match.')
                }
            }
        }
        
    }
}

