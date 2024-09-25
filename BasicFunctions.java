package org.common;

import java.awt.AWTException;
import java.awt.Robot;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.concurrent.TimeUnit;
import java.util.function.Function;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.common.BasicFunctions.DBCredentials;
import org.openqa.selenium.Alert;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;


public class BasicFunctions {
	
	public static WebDriver driver;
	public static Actions act;
	public static FileInputStream fis = null;
	public static FileOutputStream fos = null;
	public static XSSFWorkbook workbook = null;
	public static XSSFSheet sheet = null;
	public static XSSFRow row = null;
	public static XSSFCell cell = null;
	
	public static void textSend(WebElement refName, String textValue) {
		refName.sendKeys(textValue);
	}

	public static void buttonClick(WebElement refName) {
		refName.click();

	}
	public static String getTitle() {
		String title = driver.getTitle();
		System.out.println(title);
		return title;
	}
	public static String getAtrributeValue(WebElement refName, String AttributeValue) {
		String attribute = refName.getAttribute(AttributeValue);
		return attribute;

	}
	public static void currentUrl() {
		String currentUrl = driver.getCurrentUrl();
		System.out.println(currentUrl);
	}
	
	public static void implicityWait() {
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
	}
	public static void navigateTo(String url) {
		driver.navigate().to(url);

	}

	public static void navigateBack() {
		driver.navigate().back();

	}

	public static void navigateForward() {
		driver.navigate().forward();

	}

	public static void navigateRefresh() {
		driver.navigate().refresh();

	}
	public static void mouseOver(WebElement refName) {
		act = new Actions(driver);
		act.moveToElement(refName).perform();
	}

	public static void rightClick(WebElement refName) {
		act = new Actions(driver);
		act.contextClick(refName).perform();

	}

	public static void doubleClick(WebElement refName) {
		act = new Actions(driver);
		act.doubleClick(refName).perform();

	}

	public static void dragAndDrop(WebElement refNameSource, WebElement refNameTarget) {
		act = new Actions(driver);
		act.dragAndDrop(refNameSource, refNameTarget).perform();

	}

	public static void keyPress(int keycode) throws AWTException {
		Robot r = new Robot();
		r.keyPress(keycode);

	}
	public static void keyRelease(int keycode) throws AWTException {
		Robot r = new Robot();
		r.keyRelease(keycode);

	}
	public static void javaScribtClick(WebElement element) {
		JavascriptExecutor jk = (JavascriptExecutor) driver;
		jk.executeScript("arguments[0].click()", element);
	}
	public static void textSendJs(WebElement element, String data) {
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("arguments[0].setAttribute('value','" + data + "')", element);
		element.sendKeys(Keys.TAB);
	}
	
	// getAttribute Using JavaScript
	public static void getAttributeJavaScript(WebElement refName) {
		JavascriptExecutor jk = (JavascriptExecutor) driver;
		Object executeScript = jk.executeScript("return arguments[0].getttribute('value')", refName);
		System.out.println(executeScript);
	}
	// scroll Down using JavaScript
	public static void scrollDownJavaSc(WebElement element) {
		JavascriptExecutor jk = (JavascriptExecutor) driver;
		jk.executeScript("arguments[0].scrollIntoView(true);", element);
	}

	// scroll Up using JavaScript
	public static void scrollUpJavaSc(WebElement element) {
		JavascriptExecutor jk = (JavascriptExecutor) driver;
		jk.executeScript("arguments[0].scrollIntoView(false);", element);
	}
	public static void acceptAlert() {
		driver.switchTo().alert().accept();
	}

	public static void dismissAlert() {
		driver.switchTo().alert().dismiss();

	}
	
	public static String Extract_DOB(String s) {
	       

        // Define a regular expression pattern to match the date
        String dobPattern = "(\\d{2}/\\d{2}/\\d{4})";

        // Create a Pattern object
        Pattern pattern = Pattern.compile(dobPattern);

        // Create a Matcher object
        Matcher matcher = pattern.matcher(s);
      
        // Check if the pattern is found
        if (matcher.find()) {
            // Extract the matched date
        	return matcher.group(1);
        } else {
        	return null;
        }
		
    }

	public static void promptAlert(String value) {
		Alert text = driver.switchTo().alert();
		text.sendKeys(value);
		text.accept();
	}
	
	public static void alertGetText() {
		Alert alert = driver.switchTo().alert();
		String text = alert.getText();
		System.out.println(text);

	}
	public static void selectByIndex(WebElement refName, int string) {
		Select select = new Select(refName);
		select.selectByIndex(string);
	}
	public static void selectByValue(WebElement refName, String value) {
		Select select = new Select(refName);
		select.selectByValue(value);
	}
	public static void selectByVisibleText(WebElement refName, String text) {
		Select select = new Select(refName);
		select.selectByVisibleText(text);
	}
	public static String first_Selected_Value(WebElement refName) {
		Select select = new Select(refName);
		WebElement selectedOption = select.getFirstSelectedOption();
		String selectedOptionText = selectedOption.getText();
		return selectedOptionText;
	}
	public static void screenshot(String location) throws IOException {
		TakesScreenshot tks = (TakesScreenshot) driver;
		File defaultLocation = tks.getScreenshotAs(OutputType.FILE);
		System.out.println(defaultLocation);
		FileUtils.copyFile(defaultLocation, new File(location));
	}
	
	public static double string_To_double_Convert(String str_Value) {
		String replace = str_Value.replace(",", "");
		double parseDouble = Double.parseDouble(replace);
		return parseDouble;
	}
	

	public static void webDriverWait(ExpectedCondition<WebElement> ExpectedConditions) {
		WebDriverWait wait = new WebDriverWait(driver, 10);
		wait.until(ExpectedConditions);
	}

	public static void fluentWait(Class<? extends Throwable> exceptionType, Function<? super WebDriver, Object> ExpectedConditions) {
		FluentWait<WebDriver> wait = new FluentWait<WebDriver>(driver).withTimeout(10, TimeUnit.SECONDS).
				pollingEvery(5, TimeUnit.SECONDS).ignoring(exceptionType);
		wait.until(ExpectedConditions);
	}
	
	public static String  get_DB_Data(String query,String Col_Name) throws IOException, ClassNotFoundException {
		String data=null;
		 List<DBCredentials> dbCredentialsList = new ArrayList<>();

	        // Load the Excel file
	        File file = new File(System.getProperty("user.dir") + "\\ExcelUtils\\Test_Data.xlsx");
	        FileInputStream fileInputStream = new FileInputStream(file);
	        Workbook book = new XSSFWorkbook(fileInputStream);
	        Sheet sheet = book.getSheetAt(0); 

	        // Read the data
	        for (Row row : sheet) {
	            Cell flagCell = row.getCell(2); 
	            if (flagCell != null && "Y".equalsIgnoreCase(flagCell.getStringCellValue())) {
	                
	                String dbUrl = row.getCell(3).getStringCellValue();
	                String dbUser = row.getCell(4).getStringCellValue();
	                String dbPassword = row.getCell(5).getStringCellValue();
	                dbCredentialsList.add(new DBCredentials(dbUrl, dbUser, dbPassword));
	            }
	        }
	        fileInputStream.close();
	        Class.forName("oracle.jdbc.driver.OracleDriver");
	        // Connect to each database
	        for (DBCredentials creds : dbCredentialsList) {
	            try (Connection connection = DriverManager.getConnection(creds.getUrl(), creds.getUser(), creds.getPassword())) {
	                System.out.println("Connected to the database successfully: " + creds.getUrl());
	                
	                // Execute the query
	                try (PreparedStatement prepareStatement = connection.prepareStatement(query);
	                     ResultSet executeQuery = prepareStatement.executeQuery()) {
	                    while (executeQuery.next()) {
	                        data = executeQuery.getString(Col_Name);
	                       
	                      //  System.out.println(string);
	                    }
	                }
	            } catch (Exception e) {
	                System.out.println("Connection failed: " + e.getMessage());
	            }
	        }
	        return data;
	    }

	    // Helper class to hold DB credentials
	    static class DBCredentials {
	        private String url;
	        private String user;
	        private String password;

	        public DBCredentials(String url, String user, String password) {
	            this.url = url;
	            this.user = user;
	            this.password = password;
	        }

	        public String getUrl() {
	            return url;
	        }

	        public String getUser() {
	            return user;
	        }

	        public String getPassword() {
	            return password;
	        }
	}
	public static void SetStatus(String sheetName, int rowNum, String colName, String value) throws IOException  {
		fis = new FileInputStream(System.getProperty("user.dir")+"\\ExcelUtils\\Test_Data_Report.xlsx");
		workbook = new XSSFWorkbook(fis);
		fis.close();
		int colNum = -1;
		sheet = workbook.getSheet(sheetName);
		row = sheet.getRow(0);
		for (int i=1; i<row.getLastCellNum(); i++) {
			if(row.getCell(i).getStringCellValue().trim().equals(colName)) {
				colNum = i;
			}
		}
		row = sheet.getRow(rowNum);
		if(row == null)
			row = sheet.createRow(rowNum);
		
		cell = row.getCell(colNum);
				if(cell == null)
					cell = row.createCell(colNum);
				cell.setCellValue(value);
				fos = new FileOutputStream(System.getProperty("user.dir")+"\\ExcelUtils\\Test_Data_Report.xlsx");
				workbook.write(fos);
				fos.close();
	}
	
	public static void update_Status(String sheetName, String colName, List<String> testDataIdentifiers) throws IOException {
		fis = new FileInputStream(System.getProperty("user.dir")+"\\ExcelUtils\\Test_Data_Report.xlsx");
		workbook = new XSSFWorkbook(fis);
		fis.close();
		sheet = workbook.getSheet(sheetName);
		int colNum = -1;
		Row headerRow = sheet.getRow(0); // Assuming the header is in the first row
	    for (int i = 0; i < headerRow.getLastCellNum(); i++) {
	        if (headerRow.getCell(i).getStringCellValue().trim().equals(colName)) {
	            colNum = i;
	            break;
	        }
	    }

	    if (colNum == -1) {
	        System.out.println("Column not found: " + colName);
	        return;
	    }

	    if (colNum == -1) {
	        System.out.println("Column not found: " + colName);
	        return;
	    }

	    int rowNum = 1; // Assuming data starts from the second row
	    for (String testDataIdentifier : testDataIdentifiers) {
	        // Loop through each row and update the status for the specific test data identifier
	        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
	            row = sheet.getRow(i);
	            Cell testDataIdentifierCell = row.getCell(1); // Assuming the test data identifier is in the first column
	            String testDataIdentifierCellValue = testDataIdentifierCell.getStringCellValue().trim();

	            if (testDataIdentifierCellValue.equals(testDataIdentifier)) {
	                cell = row.createCell(colNum);
	                cell.setCellValue("OutputValueFor" + testDataIdentifier); // Replace with your actual output value
	                rowNum++;
	                break; // Exit the loop once the status is updated for the specific test data identifier
	            }
	        }
	    }



        fos = new FileOutputStream(System.getProperty("user.dir") + "\\ExcelUtils\\Test_Data_Report.xlsx");
        workbook.write(fos);
        fos.close();
	}
	
	public static void update_Status(String sheetName, String colName, String value, String testDataIdentifier) throws IOException {
		fis = new FileInputStream(System.getProperty("user.dir")+"\\ExcelUtils\\Test_Data_Report.xlsx");
		workbook = new XSSFWorkbook(fis);
		fis.close();
		sheet = workbook.getSheet(sheetName);
		int colNum = -1;
		Row headerRow = sheet.getRow(0); // Assuming the header is in the first row
	    for (int i = 0; i < headerRow.getLastCellNum(); i++) {
	        if (headerRow.getCell(i).getStringCellValue().trim().equals(colName)) {
	            colNum = i;
	            break;
	        }
	    }

	    if (colNum == -1) {
	        System.out.println("Column not found: " + colName);
	        return;
	    }

	    // Loop through each row and update the status for the specific test data identifier
	 //   int rownum =1;
	    int rowIndex = -1;
	    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
	        row = sheet.getRow(i);
	        Cell testDataIdentifierCell = row.getCell(0); // Assuming the test data identifier is in the first column
	     // Check if the cell type is numeric
	        if (testDataIdentifierCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
	            // Convert the numeric value to a string for comparison
	            String testDataIdentifierCellValue = String.valueOf((int) testDataIdentifierCell.getNumericCellValue());
	            if (testDataIdentifierCellValue.equals(testDataIdentifier)) {
	                cell = row.createCell(colNum);
	                cell.setCellValue(value);
	                rowIndex = i;
	                break;
	            }
	        } else if (testDataIdentifierCell.getCellType() == Cell.CELL_TYPE_STRING) {
	            // If the cell type is already a string, use getStringCellValue()
	            String testDataIdentifierCellValue = testDataIdentifierCell.getStringCellValue().trim();
	            if (testDataIdentifierCellValue.equals(testDataIdentifier)) {
	                cell = row.createCell(colNum);
	                cell.setCellValue(value);
	                rowIndex = i;
	                break;}
	        }
	    }


        fos = new FileOutputStream(System.getProperty("user.dir") + "\\ExcelUtils\\Test_Data_Report.xlsx");
        workbook.write(fos);
        fos.close();
	} 
	
	public static void updateStatus(String sheetName, String colName, String value, String testDataIdentifier) throws IOException {
         fis = new FileInputStream(System.getProperty("user.dir") + "\\ExcelUtils\\Test_Data_Report.xlsx");
         workbook = new XSSFWorkbook(fis);
        fis.close();
        Sheet sheet = workbook.getSheet(sheetName);

        int colNum = -1;
        Row headerRow = sheet.getRow(0); // Assuming the header is in the first row
        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            if (headerRow.getCell(i).getStringCellValue().trim().equals(colName)) {
                colNum = i;
                break;
            }
        }

        if (colNum == -1) {
            System.out.println("Column not found: " + colName);
            fis.close();
            return;
        }

        int rowIndex = findRowIndex(sheet, testDataIdentifier);
        if (rowIndex != -1) {
            Row row = sheet.getRow(rowIndex);
            Cell cell = row.createCell(colNum);
            cell.setCellValue(value);

            FileOutputStream fos = new FileOutputStream(System.getProperty("user.dir") + "\\ExcelUtils\\Test_Data_Report.xlsx");
            workbook.write(fos);
            fos.close();

            System.out.println("Value updated successfully for test data identifier: " + testDataIdentifier);
        } else {
            System.out.println("Test data identifier not found: " + testDataIdentifier);
        }

        fis.close();
    }

    private static int findRowIndex(Sheet sheet, String testDataIdentifier) {
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            Cell testDataIdentifierCell = row.getCell(1); // Assuming the test data identifier is in the first column
            String testDataIdentifierCellValue = testDataIdentifierCell.getStringCellValue().trim();
            if (testDataIdentifierCellValue.equals(testDataIdentifier)) {
                return i;
            }
        }
        return -1; // Return -1 if the test data identifier is not found
    }
	public static String readExcel(String sheetname, int rowno, int cellno) throws IOException  {
		String data=null;
		File file = new File("C:\\Users\\vijayaragavan.h\\eclipse-workspace\\UAT_ELMO_Retail\\ExcelUtils\\Test Case.xlsx");
		FileInputStream stream = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet(sheetname);
		Row row = sheet.getRow(rowno);
		Cell cell = row.getCell(cellno);
		int type = cell.getCellType();
		if (type == 1) {
			data = cell.getStringCellValue();
		}
		else if (DateUtil.isCellDateFormatted(cell)) {
			Date dateCellValue = cell.getDateCellValue();
			SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
			data= dateFormat.format(dateCellValue);
		} else {
			double d = cell.getNumericCellValue();
			long l = (long) d;
			data = String.valueOf(l);

		}
		return data;
	}
	
}
