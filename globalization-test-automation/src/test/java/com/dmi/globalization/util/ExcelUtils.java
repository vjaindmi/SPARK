package com.dmi.globalization.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import com.dmi.globalization.setup.Constants;
import com.dmi.globalization.util.ExcelUtils;

import io.appium.java_client.ios.IOSDriver;

@SuppressWarnings({"deprecation", "unused"})
public class ExcelUtils 
{
	public XSSFWorkbook workbook;
	public XSSFSheet sheet;
	public XSSFCell cell;
	public  XSSFRow row;
	private  FileInputStream in;

	public void setExcelFile(String path, String file) throws IOException
	{
		FileInputStream fs = new FileInputStream(path.concat(file));
		workbook = new XSSFWorkbook(fs);
	}

	public void deleteExcelSheet(String path, String file, String sSheetName) throws IOException
	{
		FileInputStream fs = new FileInputStream(path.concat(file));
		workbook = new XSSFWorkbook(fs);
		XSSFSheet sheet = workbook.getSheet(sSheetName);
		int index=0;
		if(isSheetExists(sSheetName))  
		{
			index = workbook.getSheetIndex(sheet);
			workbook.removeSheetAt(index);
			FileOutputStream output = new FileOutputStream(path.concat(file));
			workbook.write(output);
			output.close();
		}
	}

	public void deleteFile(String FileName)
	{
		File videoArchive=new File(FileName);
		try
		{
			File[] listOfFiles = videoArchive.listFiles();
			for (int file_count = 0; file_count < listOfFiles.length; file_count++) 
			{
				if((listOfFiles[file_count].getName().endsWith(".html"))||(listOfFiles[file_count].getName().endsWith(".xlsx")))
				{
					listOfFiles[file_count].delete();
				}
			}
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
	}

	public void copyFile(String src,String dest)
	{
		File source = new File(src);
		File destination = new File(dest);
		try
		{
			FileUtils.copyDirectory(source, destination);
		} 
		catch (IOException e)
		{
			e.printStackTrace();
		}
	}

	public void copyFileDir(String src,String dest)
	{
		File source = new File(src);
		File destination = new File(dest);
		try
		{
			FileUtils.copyDirectory(source, destination);
		} 
		catch (IOException e)
		{
			e.printStackTrace();
		}
	}
	
	public int counttestCases(String sheetName) throws InterruptedException
	{
		XSSFSheet sheet;
		sheet = workbook.getSheet(sheetName);
		String cellData = new String();

		int i;
		int totalrow = getRowCount(sheetName);
		int count=0;

		Thread.sleep(500);
		for (i = 1; i <=totalrow; i++)
		{
			row = sheet.getRow(i);
			cell = row.getCell(0);
			cellData = cell.getStringCellValue();
			if (cellData.equalsIgnoreCase("End of Test Cases"))
			{
				break;
			}
			else
				count++;
		}
		return count;
	}

	public int[] StartEndRows(String sheetName, String testCase, int startrow) throws InterruptedException
	{
		int startendrow[] = new int[2];
		int totalrow = getRowCount(sheetName);

		for(int i=1; i<totalrow ;i++){

			if(getCellData(sheetName, i, 0).contains(testCase) && getCellData(sheetName, i, 0).contains("Start") ){
				startendrow[0]=i;
			}
			else if(getCellData(sheetName, i, 0).contains(testCase) && getCellData(sheetName, i, 0).contains("End") ){
				startendrow[1]=i;
				break;
			}
		}

		return startendrow;
	}

	public int lastrow(String sheetName, String Test_Name, int rowNum, int colNum) throws InterruptedException
	{
		sheet = workbook.getSheet(sheetName);
		String cellData = new String();

		int i;
		int lastRowOfTC=rowNum;
		int totalrow = getRowCount(sheetName);

		Thread.sleep(500);
		for (i = rowNum; i <=totalrow; i++)
		{
			row = sheet.getRow(i);
			cell = row.getCell(colNum);
			cellData = cell.getStringCellValue();
			if (cellData.equalsIgnoreCase(Test_Name))
			{
				lastRowOfTC++;
			}
			else
				break;
		}
		return lastRowOfTC;
	}
	/* this method gets the value of the specified cell.
	 * takes sheetname, row number and column number in respective order
	 * returns the string type value of the cell
	 */
	public String getCellDatastring(String sheetName, int rowNum, int colNum) 
	{
		sheet = workbook.getSheet(sheetName);
		row = sheet.getRow(rowNum);
		cell = row.getCell(colNum);
		String cellData = new String();
		cellData = cell.getStringCellValue();
		return cellData;
	}


	@SuppressWarnings("rawtypes")
	public String mobileLocalization(IOSDriver driver) throws IOException, InterruptedException, ParserConfigurationException, SAXException 
	{
		String pagesource = driver.getPageSource();
		String text=null;

		Thread.sleep(2000);

		File file1 = new File("/Users/varunmalik/workspace/globalization-test-automation/xml/temp.txt");
		FileWriter fileWriter = new FileWriter(file1);
		pagesource = driver.getPageSource();
		fileWriter.write(pagesource);
		fileWriter.flush();
		fileWriter.close();

		DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
		DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
		Document doc = dBuilder.parse("/Users/varunmalik/workspace/globalization-test-automation/xml/temp.txt");

		doc.getDocumentElement().normalize();

		NodeList nList = doc.getElementsByTagName("XCUIElementTypeStaticText");

		for (int temp = 0; temp < nList.getLength(); temp++)
		{
			Node nNode = nList.item(temp);

			if (nNode.getNodeType() == Node.ELEMENT_NODE)
			{
				Element eElement = (Element) nNode;
				System.out.println(eElement.getAttribute("value"));
				text+=eElement.getAttribute("value")+" ";
			}
		}

		nList = doc.getElementsByTagName("XCUIElementTypeOther"); 

		for (int temp = 0; temp < nList.getLength(); temp++) 
		{
			Node nNode = nList.item(temp);

			if (nNode.getNodeType() == Node.ELEMENT_NODE) 
			{
				Element eElement = (Element) nNode;
				if(!eElement.getAttribute("label").isEmpty())
				{
					System.out.println(eElement.getAttribute("label"));
					text+=eElement.getAttribute("label")+" ";
				}
			}
		}

		nList = doc.getElementsByTagName("XCUIElementTypeButton");

		for (int temp = 0; temp < nList.getLength(); temp++) 
		{
			Node nNode = nList.item(temp);

			if (nNode.getNodeType() == Node.ELEMENT_NODE) 
			{
				Element eElement = (Element) nNode;
				if(!eElement.getAttribute("label").isEmpty())
				{
					System.out.println(eElement.getAttribute("label"));
					text+=eElement.getAttribute("label")+" ";
				}
			}
		}


		return text;
	}

	public void testLocalization(ExcelUtils localExcel, String Language, String screenName, String entireVisibleText, String tabName) throws IOException, InterruptedException 
	{

		int startendrow[] = localExcel.StartEndRows(screenName, tabName, -4);

		String Word;
		int numOfRowsinLanguageSheet = localExcel.getRowCount(screenName);

		for(int j=startendrow[0]; j<=startendrow[1]; j++)
		{
			Word = localExcel.getCellData(screenName, j, 2);
			int count = 0;

			Pattern p = Pattern.compile(Word);
			Matcher m = p.matcher(entireVisibleText);
			while (m.find()) {
				count++;
			}

			if (count>0)
			{
				localExcel.setCellData(Language, screenName, "Present", j, 3);
				localExcel.setCellData(Language, screenName, String.valueOf(count), j, 4);
			}
			else
			{
				localExcel.setCellData(Language, screenName, "Not Found", j, 3);
				localExcel.setCellData(Language, screenName, String.valueOf(count), j, 4);
			}
		}


	}

	public String getCellData(String sheetName, int rowNum, int colNum) 
	{
		sheet = workbook.getSheet(sheetName);
		row = sheet.getRow(rowNum);
		cell = row.getCell(colNum);
		String cellData = new String();
		try
		{
			if (cell == null) 
			{
				cellData = "";
				return cellData;
			}
			else if (cell.getCellType() == Cell.CELL_TYPE_STRING) 
			{
				cellData = cell.getStringCellValue();
				return cellData;
			}
			else if (cell.getCellType() == Cell.CELL_TYPE_BLANK)
			{
				return cellData;
			} 
			else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC || cell.getCellType() == Cell.CELL_TYPE_FORMULA)
			{
				cellData = String.valueOf(cell.getNumericCellValue());
			}

			return cellData;
		} 
		catch (Exception e)
		{
			e.printStackTrace();
			return "";
		}
	}

	/* this method gets the row number for a specified cell value
	 * takes input as the sheet name, string value and column number in respective order.
	 * it compares the cell values in a loop with the given string value to get the row number of that cell.
	 * integer value of row number is returned
	 */
	public int rowNum(String sheetName, String testCase, int colNum)
	{
		int i;
		for (i = 0; i <= getRowCount(sheetName); i++)
		{
			if (getCellData(sheetName, i, colNum).equalsIgnoreCase(testCase))
			{
				break;
			}
		}
		return i;
	}

	/* this method gets the total number of rows in a sheet
	 * takes the input as sheetname
	 * returns an integer value
	 */
	public int getRowCount(String sheetName)
	{
		sheet = workbook.getSheet(sheetName);
		int number = sheet.getLastRowNum();
		return number+1;
	}

	public Sheet currentSheet(String sheetName)
	{
		sheet = workbook.getSheet(sheetName);
		return sheet;
	}

	public int getColCount(String sheetName)
	{
		sheet = workbook.getSheet(sheetName);
		Row r = sheet.getRow(0);
		int number = r.getLastCellNum();
		return number;
	}

	/* this method gets the number of rows 
	 * used to get the number of test iterations in test data for a particular test case
	 */
	public int getRowLength(String sheetName, String testCase)
	{
		sheet = workbook.getSheet(sheetName);
		int j = 0,i;
		int k = getRowCount(sheetName);
		for (i = 0; i <= k; i++) 
		{
			if (getCellData(sheetName,i,0).equalsIgnoreCase(testCase)) 
			{
				//i is the row number in test data sheet that contains the test case name in first column
				//two is added because test data iterations starts two columns beneath the test case name
				i+=2;
				for (j=i;;j++)
				{     
					if(getCellData(sheetName, j, 0).equalsIgnoreCase(""))
					{
						return (j-i);
					}
				}
			}
		}
		return 0;        
	}

	public int getTestStepCount(String sheetName, String testCase, int colNum)
	{
		int number;
		for(number =0; number<=getRowCount(sheetName); number++)
		{
			if(getCellData(sheetName, number, colNum).equalsIgnoreCase(testCase))
			{
				break;
			}
		}

		return number+1;
	}

	/* this method sets a specified value to the cell.
	 * takes sheetname, string value, row number and column number as input in the respective order.
	 * Other cell formatting are also being done.
	 */
	public void setCellData(String Excel, String sheetName, String sResult, int rowNum, int colNum) throws IOException
	{
		String excelName=Excel+".xlsx";
		sheet = workbook.getSheet(sheetName);
		row = sheet.getRow(rowNum);
		cell = row.createCell(colNum);
		cell.setCellValue("");
		cell = row.getCell(colNum);
		XSSFCellStyle style = workbook.createCellStyle();
		//set color
		if(sResult=="Not Found")
		{
			style.setFillForegroundColor((new XSSFColor(new java.awt.Color(176,23,31))));
			style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		}
		else if(sResult=="Present")
		{
			style.setFillForegroundColor((new XSSFColor(new java.awt.Color(50,205,50))));
			style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		}

		if(cell==null)
		{
			cell = row.createCell(colNum);
			cell.setCellValue(sResult);
		}
		else
		{
			cell.setCellValue(sResult);
		}

		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_TOP);
		style.setBorderBottom((short) 1);
		style.setBorderLeft((short) 1);
		style.setBorderRight((short) 1);
		style.setBorderTop((short) 1);
		cell.setCellStyle(style); 

		FileOutputStream fo = new FileOutputStream(Constants.sReportsPath.concat(excelName));
		workbook.write(fo);
		fo.close();
	}

	/* this method checks if a sheet exists
	 * takes input as sheet name
	 * returns boolean value 
	 */
	public boolean isSheetExists(String sheetName)
	{
		List<String> sheetNames = new ArrayList<String>();
		for (int i=0; i<workbook.getNumberOfSheets(); i++) 
		{
			sheetNames.add(workbook.getSheetName(i) );
		}

		if(sheetNames.contains(sheetName))
		{
			return true;
		}
		else
		{
			return false;
		}
	}

	// this method removes contents from the specified cells.
	public boolean removeContent(String Path, String fileName, String sheetName, int colNum)
	{
		try
		{
			if(!isSheetExists(sheetName))
				return false;
			String file = Path.concat(fileName); 
			FileInputStream fs = new FileInputStream(file);
			workbook = new XSSFWorkbook(fs);
			fs = new FileInputStream(file); 
			workbook = new XSSFWorkbook(fs);
			sheet=workbook.getSheet(sheetName);

			for(int i=1;i<=getRowCount(sheetName);i++)
			{
				row=sheet.getRow(i); 

				if(row!=null)
				{
					cell=row.getCell(colNum);

					if(cell!=null)
					{
						row.removeCell(cell);
					}
				}
			}
			FileOutputStream fo = new FileOutputStream(file);
			workbook.write(fo);
			fo.close();
		}
		catch(Exception e)
		{
			e.printStackTrace();
			return false;
		}
		return true;
	}
}
