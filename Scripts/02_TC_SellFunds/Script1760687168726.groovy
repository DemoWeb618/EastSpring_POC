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
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.webui.driver.DriverFactory
import org.openqa.selenium.WebDriver

WebUI.openBrowser('')

WebUI.navigateToUrl('https://services.uat.eastspring.co.th/agent-ie-decom/')

WebUI.setText(findTestObject('Object Repository/Page_TA-SYSTEM  Eastspring Asset Management/input_Login_login_name'), 'gable_4')

WebUI.setEncryptedText(findTestObject('Object Repository/Page_TA-SYSTEM  Eastspring Asset Management/input_Password_password'), 
    'XhlV0WqOzrTq9A21Os2qrA==')

WebUI.click(findTestObject('Object Repository/Page_TA-SYSTEM  Eastspring Asset Management/input_Password_submit'))

WebUI.click(findTestObject('Object Repository/Page_/a__hyperlink_4'))

WebUI.setText(findTestObject('Object Repository/Page_/input__holder_no'), acc_no)

WebUI.selectOptionByValue(findTestObject('Object Repository/Page_/select__tmp_fund_code'), funds_id, true)

WebUI.setText(findTestObject('Object Repository/Page_/input__redeem_amount_in_baht'), amount)

WebUI.click(findTestObject('Object Repository/Page_/input_Note_submit'))

WebUI.waitForAlert(5)

WebUI.acceptAlert()

WebUI.click(findTestObject('Object Repository/Page_/input_Local AuthorizeRemote Authorize_choose_auth'))

WebUI.setText(findTestObject('Object Repository/Page_/input_Supervisor ID_teller_login'), 'gable_5')

WebUI.setEncryptedText(findTestObject('Object Repository/Page_/input_Password_teller_password'), 'rCOjHjf4WNFiBYyPgzDSng==')

WebUI.click(findTestObject('Object Repository/Page_/input_Password_fill_form2'))

WebUI.delay(10)
// Get the current WebDriver instance
WebDriver driver = DriverFactory.getWebDriver()

// Get all window handles (each open tab/window)
Set<String> allWindows = driver.getWindowHandles()

for (String window : allWindows) {
	driver.switchTo().window(window)
	String currentUrl = driver.getCurrentUrl()

	if (currentUrl.contains('agent-ie-decom/sa/') && currentUrl.contains('/voucher.jsp')) {
		println("✅ Switched to correct window: " + currentUrl)
		break
	}
}

WebUI.waitForElementPresent(findTestObject('Object Repository/Page_/u'), 0)

WebUI.verifyElementText(findTestObject('Object Repository/Page_/td'), acc_no)

WebUI.verifyElementText(findTestObject('Object Repository/Page_/td_1'), 'คุณA05 0000000012')

WebUI.verifyElementText(findTestObject('Object Repository/Page_/td_6'), funds_fullname)

WebUI.verifyElementText(findTestObject('Object Repository/Page_/td_7'), '50.00 บาท')

//WebUI.verifyElementText(findTestObject('Object Repository/Page_/td_4'), '17/10/202517:55 น.')

WebUI.switchToWindowUrl('https://services.uat.eastspring.co.th/agent-ie-decom/login/login_accept.jsp')

WebUI.click(findTestObject('Object Repository/Page_/b'))

WebUI.closeBrowser()

