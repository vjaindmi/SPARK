package com.dmi.globalization.execution;

import java.io.IOException;
import java.net.URL;
import java.util.HashMap;
import java.util.Map;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.testng.annotations.Test;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import com.dmi.globalization.setup.Constants;
import com.dmi.globalization.util.ExcelUtils;

import io.appium.java_client.ios.IOSDriver;


public class Driver
{
	@SuppressWarnings("rawtypes")

	@Test
	public static void main() throws InterruptedException, IOException, ParserConfigurationException, SAXException
	{
		int numOfRows;
		String runMode;
		String language;
		com.dmi.globalization.util.ExcelUtils mainExcel = new ExcelUtils();
		String entireVisibleText;
		Map<String, String> myMap = new HashMap<String, String>();

		mainExcel.setExcelFile(Constants.sTestDataPath, Constants.DataEngine);
		mainExcel.deleteFile(Constants.Srcfolder);
		mainExcel.copyFileDir(Constants.sTestDataPath, Constants.Srcfolder);
		numOfRows = mainExcel.getRowCount(Constants.Driver);

		IOSDriver driver;

		for(int i=2; i<numOfRows; i++){
			runMode = mainExcel.getCellData(Constants.Driver, i, 3);

			if (runMode.equalsIgnoreCase("Y"))
			{
				language = mainExcel.getCellData(Constants.Driver, i, 2);

				try
				{
					System.out.println(language);
					DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
					DocumentBuilder db = dbf.newDocumentBuilder();
					Document doc = db.parse("/Users/varunmalik/workspace/globalization-automation/xml/"+language+"_strings.xml");

					NodeList nodes = doc.getElementsByTagName("string");

					for (int j = 0; j < nodes.getLength(); j++)
					{
						Node nNode =  nodes.item(j);
						Element eElement = (Element) nNode;
						myMap.put(eElement.getAttribute("name"), nodes.item(j).getTextContent());
					}
				}
				catch (Exception e)
				{
					e.printStackTrace();
				}

				DesiredCapabilities capabilities = new DesiredCapabilities();

				capabilities.setCapability("platformName", "iOS");
				capabilities.setCapability("noReset", "true");
				//capabilities.setCapability("deviceName", "LB iphone 6s");
				capabilities.setCapability("deviceName", "iPhone 7");
				capabilities.setCapability("platformVersion", "11.2");
				capabilities.setCapability("automationName", "XCUITest");
				capabilities.setCapability("app", System.getProperty("user.dir")+"/apps/iphone-apps/simulator-app/OpenLink.app");
				//capabilities.setCapability("udid", "3056cedf0f3674c10b594d956560726f91feb7b6");
				//capabilities.setCapability("udid", "049b5c278059ea392be55b49e26cbbc3317d6682");
				capabilities.setCapability("bundleId","com.apple.Preferences");
				capabilities.setCapability("newCommandTimeout", 10000);
				capabilities.setCapability("noReset", "true");

				driver = new IOSDriver(new URL("http://0.0.0.0:4723/wd/hub"), capabilities);
				driver.closeApp();
				
				driver = new IOSDriver(new URL("http://0.0.0.0:4723/wd/hub"), capabilities);

				System.out.println("Driver created");
				Thread.sleep(1000);

				JavascriptExecutor js = (JavascriptExecutor) driver;
				HashMap<String, String> scrollObject = new HashMap<String, String>();
				scrollObject.put("direction", "down");

				if(language.equalsIgnoreCase("French"))
				{
					//js.executeScript("mobile: scroll", scrollObject);

					driver.findElement(By.name("General")).click();

					//js.executeScript("mobile: scroll", scrollObject);

					driver.findElement(By.name("Language & Region")).click();
					driver.findElement(By.name("iPhone Language")).click();
					driver.findElement(By.name("Search")).sendKeys("French (Canada)");
					driver.findElement(By.name("French (Canada)")).click();
					driver.findElement(By.name("Done")).click();
					driver.findElement(By.name("Change to French (Canada)")).click();

				}
				else if(language.equalsIgnoreCase("English"))
				{
					//js.executeScript("mobile: scroll", scrollObject);
					
					driver.findElement(By.name("Général")).click();

					//js.executeScript("mobile: scroll", scrollObject);

					driver.findElement(By.name("Langue et région")).click();
					driver.findElement(By.name("Langue de l’iPhone")).click();
					driver.findElement(By.name("Recherche")).sendKeys("English (U.S.)");
					driver.findElement(By.name("English (U.S.)")).click();
					driver.findElement(By.name("OK")).click();
					driver.findElement(By.name("Passer en anglais (É.-U.)")).click();
				}

				Thread.sleep(24000);
				driver.closeApp();

				capabilities.setCapability("bundleId","com.milwaukeetool.mymilwaukee");
				driver = new IOSDriver(new URL("http://0.0.0.0:4723/wd/hub"), capabilities);

				ExcelUtils localExcel = new ExcelUtils();
				localExcel.setExcelFile(Constants.sReportsPath, language+".xlsx");

				entireVisibleText=mainExcel.mobileLocalization(driver);
				mainExcel.testLocalization(localExcel, language, "Login Screen", entireVisibleText, "First Screen");

				Thread.sleep(3000);
				driver.findElement(By.id(myMap.get("landing_screen_sign_in_text"))).click();

				entireVisibleText=mainExcel.mobileLocalization(driver);
				mainExcel.testLocalization(localExcel, language, "Login Screen", entireVisibleText, "Login Screen");

				driver.findElement(By.id(myMap.get("landing_screen_sign_in_text"))).click();
				driver.findElement(By.id(myMap.get("ts_sign_in_field_title_email_address_or_guest_username"))).sendKeys("vm@demo.com");
				driver.findElement(By.id(myMap.get("create_account_field_password"))).sendKeys("miP4cvma");
				driver.findElement(By.id(myMap.get("landing_screen_sign_in_text"))).click();

				Thread.sleep(6000);

				entireVisibleText=mainExcel.mobileLocalization(driver);
				mainExcel.testLocalization(localExcel, language, "Manage Inventory", entireVisibleText, "Inventory");

				driver.findElement(By.id(myMap.get("main_title_settings"))).click();

				entireVisibleText=mainExcel.mobileLocalization(driver);
				mainExcel.testLocalization(localExcel, language, "Settings", entireVisibleText, "Settings");

				driver.findElement(By.id(myMap.get("my_profile_title"))).click();

				entireVisibleText=mainExcel.mobileLocalization(driver);
				mainExcel.testLocalization(localExcel, language, "Settings", entireVisibleText, "My Profile");

				driver.findElement(By.id(myMap.get("dialog_save_profile_cancel"))).click();
				
				Thread.sleep(3000);
				
				js = (JavascriptExecutor) driver;
				js.executeScript("mobile: scroll", scrollObject);
				js.executeScript("mobile: scroll", scrollObject);

				driver.findElement(By.id(myMap.get("btn_title_log_out"))).click();
			}
		}
	}
}