import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testng.keyword.TestNGBuiltinKeywords as TestNGKW
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import internal.GlobalVariable as GlobalVariable
import org.openqa.selenium.Keys as Keys
import com.kms.katalon.core.webui.driver.DriverFactory as DriverFactory
import java.awt.Toolkit as Toolkit
import java.awt.datatransfer.DataFlavor as DataFlavor
import org.apache.poi.xssf.usermodel.XSSFWorkbook as XSSFWorkbook
import org.apache.poi.ss.usermodel.Row as Row
import org.apache.poi.ss.usermodel.Cell as Cell
import org.apache.poi.ss.usermodel.Sheet as Sheet
import java.io.FileOutputStream as FileOutputStream

Today = CustomKeywords.'datetime.DT.getToday_ddMMyyyy'()

WebUI.openBrowser('')

WebUI.navigateToUrl('https://services.uat.eastspring.co.th/agent-ie-decom/')

WebUI.takeScreenshot()

WebUI.setText(findTestObject('Object Repository/Page_TA-SYSTEM  Eastspring Asset Management/input_Login_login_name'), 'gable_5')

WebUI.setEncryptedText(findTestObject('Object Repository/Page_TA-SYSTEM  Eastspring Asset Management/input_Password_password'), 
    'rCOjHjf4WNFiBYyPgzDSng==')

WebUI.click(findTestObject('Object Repository/Page_TA-SYSTEM  Eastspring Asset Management/input_Password_submit'))

WebUI.click(findTestObject('Object Repository/Page_/a__hyperlink_3'))

WebUI.click(findTestObject('Object Repository/Page_/a__hyperlink_1'))

WebUI.selectOptionByValue(findTestObject('Object Repository/Page_/select__tmp_fund_code'), FundsID, true)

WebUI.click(findTestObject('Object Repository/Page_/input__btn_pur'))

WebUI.takeScreenshot()

boolean Buy_Save_to_excel_available = WebUI.waitForElementPresent(findTestObject('Object Repository/Page_/input__copy_btn'), 
    10)

if (Buy_Save_to_excel_available) {
    WebUI.click(findTestObject('Object Repository/Page_/input__copy_btn'))

    WebUI.waitForAlert(5)

    WebUI.acceptAlert()

    WebUI.switchToWindowUrl('https://services.uat.eastspring.co.th/agent-ie-decom/login/login_accept.jsp')

    println('✅ \'Save to Excel\' button found and clicked.')

    // Step 2: Get clipboard text
    def clipboard = Toolkit.getDefaultToolkit().getSystemClipboard()

    def rawData = clipboard.getData(DataFlavor.stringFlavor)

    // Print to verify
    println('📋 Clipboard raw data: ' + rawData)

    println('📋 Clipboard length: ' + rawData.length())

    // Step 3: Create new Excel workbook and sheet
    XSSFWorkbook workbook = new XSSFWorkbook()

    Sheet sheet = workbook.createSheet('Voucher Data')

    // Step 4: Split clipboard data into rows and columns
    def rows = rawData.split('\\r?\\n' // Split by newline
        )

    int rowIndex = 0

    for (String line : rows) {
        if (line.trim().isEmpty()) // skip blank lines
        {
            continue
        }
        
        Row row = sheet.createRow(rowIndex++)

        def cols = line.split('\\t' // Split by tab between cells
            )

        for (int colIndex = 0; colIndex < cols.length; colIndex++) {
            Cell cell = row.createCell(colIndex)

            cell.setCellValue(cols[colIndex])
        }
    }
    
    // Step 5: Autosize columns
    int totalCols

    for (int i = 0; i < totalCols; i++) {
        sheet.autoSizeColumn(i)
    }
    
    // Step 6: Save Excel file
    String filePath = "D:\\Project\\POC_EastSpring\\BuyTransaction_${FundsID}_${Today}.xlsx"

    FileOutputStream out = new FileOutputStream(filePath)

    workbook.write(out)

    out.close()

    workbook.close()

    println('✅ Excel file created successfully at: ' + filePath)
} else {
    println('⚠️ \'Save to Excel\' button not found — skip clicking.')
}

WebUI.click(findTestObject('Object Repository/Page_/input__btn_red'))

WebUI.takeScreenshot()

boolean Red_Save_to_excel_available = WebUI.waitForElementPresent(findTestObject('Object Repository/Page_/input__copy_btn'), 
    10)

if (Red_Save_to_excel_available) {
    WebUI.click(findTestObject('Object Repository/Page_/input__copy_btn'))

    WebUI.waitForAlert(5)

    WebUI.acceptAlert()

    WebUI.switchToWindowUrl('https://services.uat.eastspring.co.th/agent-ie-decom/login/login_accept.jsp')

    println('✅ \'Save to Excel\' button found and clicked.')

    // Step 2: Get clipboard text
    def clipboard = Toolkit.getDefaultToolkit().getSystemClipboard()

    def rawData = clipboard.getData(DataFlavor.stringFlavor)

    // Print to verify
    println('📋 Clipboard raw data: ' + rawData)

    println('📋 Clipboard length: ' + rawData.length())

    // Step 3: Create new Excel workbook and sheet
    XSSFWorkbook workbook = new XSSFWorkbook()

    Sheet sheet = workbook.createSheet('Voucher Data')

    // Step 4: Split clipboard data into rows and columns
    def rows = rawData.split('\\r?\\n' // Split by newline
        )

    int rowIndex = 0

    for (String line : rows) {
        if (line.trim().isEmpty()) // skip blank lines
        {
            continue
        }
        
        Row row = sheet.createRow(rowIndex++)

        def cols = line.split('\\t' // Split by tab between cells
            )

        for (int colIndex = 0; colIndex < cols.length; colIndex++) {
            Cell cell = row.createCell(colIndex)

            cell.setCellValue(cols[colIndex])
        }
    }
    
    // Step 5: Autosize columns
    int totalCols

    for (int i = 0; i < totalCols; i++) {
        sheet.autoSizeColumn(i)
    }
    
    // Step 6: Save Excel file
    String filePath = "D:\\Project\\POC_EastSpring\\RedTransaction_${FundsID}_${Today}.xlsx"

    FileOutputStream out = new FileOutputStream(filePath)

    workbook.write(out)

    out.close()

    workbook.close()

    println('✅ Excel file created successfully at: ' + filePath)
} else {
    println('⚠️ \'Save to Excel\' button not found — skip clicking.')
}

WebUI.click(findTestObject('Object Repository/Page_/input__btn_swi'))

WebUI.takeScreenshot()

boolean SWI_Save_to_excel_available = WebUI.waitForElementPresent(findTestObject('Object Repository/Page_/input__copy_btn'), 
    10)

if (SWI_Save_to_excel_available) {
    WebUI.click(findTestObject('Object Repository/Page_/input__copy_btn'))

    WebUI.waitForAlert(5)

    WebUI.acceptAlert()

    WebUI.switchToWindowUrl('https://services.uat.eastspring.co.th/agent-ie-decom/login/login_accept.jsp')

    println('✅ \'Save to Excel\' button found and clicked.')

    // Step 2: Get clipboard text
    def clipboard = Toolkit.getDefaultToolkit().getSystemClipboard()

    def rawData = clipboard.getData(DataFlavor.stringFlavor)

    // Print to verify
    println('📋 Clipboard raw data: ' + rawData)

    println('📋 Clipboard length: ' + rawData.length())

    // Step 3: Create new Excel workbook and sheet
    XSSFWorkbook workbook = new XSSFWorkbook()

    Sheet sheet = workbook.createSheet('Voucher Data')

    // Step 4: Split clipboard data into rows and columns
    def rows = rawData.split('\\r?\\n' // Split by newline
        )

    int rowIndex = 0

    for (String line : rows) {
        if (line.trim().isEmpty()) // skip blank lines
        {
            continue
        }
        
        Row row = sheet.createRow(rowIndex++)

        def cols = line.split('\\t' // Split by tab between cells
            )

        for (int colIndex = 0; colIndex < cols.length; colIndex++) {
            Cell cell = row.createCell(colIndex)

            cell.setCellValue(cols[colIndex])
        }
    }
    
    // Step 5: Autosize columns
    int totalCols

    for (int i = 0; i < totalCols; i++) {
        sheet.autoSizeColumn(i)
    }
    
    // Step 6: Save Excel file
    String filePath = "D:\\Project\\POC_EastSpring\\SWITransaction_${FundsID}_${Today}.xlsx"

    FileOutputStream out = new FileOutputStream(filePath)

    workbook.write(out)

    out.close()

    workbook.close()

    println('✅ Excel file created successfully at: ' + filePath)
} else {
    println('⚠️ \'Save to Excel\' button not found — skip clicking.')
}

WebUI.click(findTestObject('Object Repository/Page_/input__btn_swo'))

WebUI.takeScreenshot()

boolean SWO_Save_to_excel_available = WebUI.waitForElementPresent(findTestObject('Object Repository/Page_/input__copy_btn'), 
    10)

if (SWO_Save_to_excel_available) {
    WebUI.click(findTestObject('Object Repository/Page_/input__copy_btn'))

    WebUI.waitForAlert(5)

    WebUI.acceptAlert()

    WebUI.switchToWindowUrl('https://services.uat.eastspring.co.th/agent-ie-decom/login/login_accept.jsp')

    println('✅ \'Save to Excel\' button found and clicked.')

    // Step 2: Get clipboard text
    def clipboard = Toolkit.getDefaultToolkit().getSystemClipboard()

    def rawData = clipboard.getData(DataFlavor.stringFlavor)

    // Print to verify
    println('📋 Clipboard raw data: ' + rawData)

    println('📋 Clipboard length: ' + rawData.length())

    // Step 3: Create new Excel workbook and sheet
    XSSFWorkbook workbook = new XSSFWorkbook()

    Sheet sheet = workbook.createSheet('Voucher Data')

    // Step 4: Split clipboard data into rows and columns
    def rows = rawData.split('\\r?\\n' // Split by newline
        )

    int rowIndex = 0

    for (String line : rows) {
        if (line.trim().isEmpty()) // skip blank lines
        {
            continue
        }
        
        Row row = sheet.createRow(rowIndex++)

        def cols = line.split('\\t' // Split by tab between cells
            )

        for (int colIndex = 0; colIndex < cols.length; colIndex++) {
            Cell cell = row.createCell(colIndex)

            cell.setCellValue(cols[colIndex])
        }
    }
    
    // Step 5: Autosize columns
    int totalCols

    for (int i = 0; i < totalCols; i++) {
        sheet.autoSizeColumn(i)
    }
    
    // Step 6: Save Excel file
    String filePath = "D:\\Project\\POC_EastSpring\\SWOTransaction_${FundsID}_${Today}.xlsx"

    FileOutputStream out = new FileOutputStream(filePath)

    workbook.write(out)

    out.close()

    workbook.close()

    println('✅ Excel file created successfully at: ' + filePath)
} else {
    println('⚠️ \'Save to Excel\' button not found — skip clicking.')
}

WebUI.click(findTestObject('Object Repository/Page_/input_20_cancel'))

WebUI.click(findTestObject('Object Repository/Page_/b'))

WebUI.takeScreenshot()

WebUI.closeBrowser()

