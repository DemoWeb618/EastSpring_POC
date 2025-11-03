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
import com.kms.katalon.core.webui.keyword.internal.WebUIAbstractKeyword
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import internal.GlobalVariable as GlobalVariable
import org.openqa.selenium.Keys as Keys
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.webui.driver.DriverFactory
import org.openqa.selenium.WebDriver

WebUI.openBrowser('')

WebUI.navigateToUrl(GlobalVariable.G_URL)

WebUI.takeScreenshot()

WebUI.setText(findTestObject('Object Repository/Page_TA-SYSTEM  Eastspring Asset Management/input_Login_login_name'), 'gable_4')

WebUI.setEncryptedText(findTestObject('Object Repository/Page_TA-SYSTEM  Eastspring Asset Management/input_Password_password'), 
    'XhlV0WqOzrTq9A21Os2qrA==')

WebUI.click(findTestObject('Object Repository/Page_TA-SYSTEM  Eastspring Asset Management/input_Password_submit'))

WebUI.click(findTestObject('Object Repository/Page_/a__hyperlink'))

WebUI.setText(findTestObject('Object Repository/Page_/input__holder_no'), acc_no)

WebUI.selectOptionByValue(findTestObject('Object Repository/Page_/select__tmp_fund_code'), funds_id, true)

WebUI.selectOptionByValue(findTestObject('Object Repository/Page_/select__tmp_tran_subtype'), payment_type, true)

WebUI.setText(findTestObject('Object Repository/Page_/input__order_amt'), amount)

WebUI.selectOptionByValue(findTestObject('Object Repository/Page_/select_Subscription Account_tmp_subscription_info'), '002:031-3-00485-5', 
    true)

WebUI.click(findTestObject('Object Repository/Page_/input__c_accept_fund_risk'))

WebUI.click(findTestObject('Object Repository/Page_/input__c_accept_fx_risk'))

WebUI.click(findTestObject('Object Repository/Page_/input__tmp_accept_conc'))

//WebUI.takeScreenshot('D:/Project/POC_EastSpring/Automate_script/PoC_EastSpring/Screenshot/BeforeSubmit.png')

WebUI.takeScreenshot()

WebUI.click(findTestObject('Object Repository/Page_/input_(High Net Worth)_submit'))

boolean buy_overtime = WebUI.waitForAlert(5)

if (buy_overtime) {
	
	WebUI.acceptAlert()
	
	WebUI.click(findTestObject('Object Repository/Page_/input_Local AuthorizeRemote Authorize_choose_auth'))
	
	WebUI.setText(findTestObject('Object Repository/Page_/input_Supervisor ID_teller_login'), 'gable_5')
	
	WebUI.setEncryptedText(findTestObject('Object Repository/Page_/input_Password_teller_password'), 'rCOjHjf4WNFiBYyPgzDSng==')
	
	WebUI.takeScreenshot()
	
	WebUI.click(findTestObject('Object Repository/Page_/input_Password_fill_form2'))
}

//Get_WD_1 = WebUI.getUrl()

//println(Get_WD_1)

//WebUI.switchToWindowTitle('https://services.uat.eastspring.co.th/agent-ie-decom/sa/202510160230/jsp/voucher.jsp?executed=YES')

// Wait a few seconds for the new tab/window to open
WebUI.delay(10)

// Get the current WebDriver instance
WebDriver driver = DriverFactory.getWebDriver()

// Get all window handles (each open tab/window)
Set<String> allWindows = driver.getWindowHandles()

// Loop through them to find the correct one
for (String window : allWindows) {
    driver.switchTo().window(window)
    String currentUrl = driver.getCurrentUrl()

    if (currentUrl.contains('agent-ie-decom/sa/') && currentUrl.contains('/voucher.jsp')) {
        println("✅ Switched to correct window: " + currentUrl)
        break
    }
}	

Get_WD_2 = WebUI.getUrl()

println(Get_WD_2)

WebUI.waitForElementPresent(findTestObject('Object Repository/Page_/u'), 0)

WebUI.verifyElementText(findTestObject('Object Repository/Page_/td'), acc_no)

WebUI.verifyElementText(findTestObject('Object Repository/Page_/td_1'), 'คุณA05 7600018928')

WebUI.verifyElementText(findTestObject('Object Repository/Page_/td_6'), funds_fullname)

//WebUI.verifyElementText(findTestObject('Object Repository/Page_/td_7'), '100.00 บาท')

//WebUI.verifyElementText(findTestObject('Object Repository/Page_/td'), '1,000.00 บาท')

//WebUI.takeScreenshot('D:/Project/POC_EastSpring/Automate_script/PoC_EastSpring/Screenshot/SubmitCompleted.png')

WebUI.takeScreenshot()

WebUI.takeScreenshotAsCheckpoint('current_viewport')

//WebUI.verifyElementText(findTestObject('Object Repository/Page_/td_1'), '17/10/202514:34 น.')

WebUI.switchToWindowUrl('https://services.uat.eastspring.co.th/agent-ie-decom/login/login_accept.jsp')

WebUI.click(findTestObject('Object Repository/Page_/b'))

WebUI.closeBrowser()

