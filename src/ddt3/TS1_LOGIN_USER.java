package ddt3;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openqa.selenium.By; 
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

class TS1_LOGIN_USER {

    @Test
    void testCheckLogin() throws IOException {
        System.setProperty("webdriver.chrome.driver", "D:/chromedriver-win64/chromedriver.exe");

        String path = "D:/A.Narupon/Data_testing/DataDrivenTesting-main/testdata1.xlsx";
        String testDate = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss").format(new Date());
        String testerName = "Suttiporn Kaewsakunnee";
        int timeoutInSeconds = 10; // Use WebDriverWait timeout

        try (FileInputStream fs = new FileInputStream(path);
             XSSFWorkbook workbook = new XSSFWorkbook(fs)) {

            XSSFSheet sheet = workbook.getSheetAt(0);
            int rowCount = sheet.getLastRowNum();
            
            WebDriver driver = new ChromeDriver();
            WebDriverWait wait = new WebDriverWait(driver, timeoutInSeconds);
            
            try {
                for (int i = 1; i <= rowCount; i++) {
                    Row currentRow = sheet.getRow(i);

                    if (currentRow == null || currentRow.getCell(0) == null || currentRow.getCell(0).toString().trim().isEmpty()) {
                        continue; // Skip empty rows
                    }

                    driver.get("http://localhost:5173/");
                    wait.until(ExpectedConditions.elementToBeClickable(By.id("GetStarted"))).click();

                    String testcaseid = currentRow.getCell(0).toString();
                    String username = testcaseid.equals("tc101") ? "" : currentRow.getCell(1).toString();
                    String password = testcaseid.equals("tc101") ? "" : currentRow.getCell(2).toString();

                    driver.findElement(By.name("email")).sendKeys(username);
                    driver.findElement(By.name("password")).sendKeys(password);
                    driver.findElement(By.id("Login")).click();

                    // Wait for the result message to be visible
                    String expectedMessage = currentRow.getCell(3) != null ? currentRow.getCell(3).toString() : "";
                    String actualMessage = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[2]/div"))).getText();

                    // Write results to Excel
                    Cell resultCell = currentRow.createCell(4);
                    resultCell.setCellValue(actualMessage);
                    Cell statusCell = currentRow.createCell(5);
                    statusCell.setCellValue(expectedMessage.equals(actualMessage) ? "Pass" : "Fail");
                    currentRow.createCell(6).setCellValue(testDate);
                    currentRow.createCell(7).setCellValue(testerName);


                    // Save Excel file
                    try (FileOutputStream fos = new FileOutputStream(path)) {
                        workbook.write(fos);
                    }
                }
            } finally {
                driver.quit(); // Ensure driver is closed
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
