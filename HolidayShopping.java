/**
 * This is a Web test automation script that automates the holiday shopping process
 * It goes to Ebay, searches for Canon camera eos 5d mark iv
 * compares prices against Excel sheet to find a good deal, then purchases it
 * <p>
 * Technical Overview:
 * Automated shopping with selenium web driver automation
 * Uses TestNG with emailable-report listener, Assert, and wait()
 * Uses data from Excel file
 * Takes screenshot throughout shopping process
 * Uses Log4j and Reporter
 * Creates TestReport at end
 * Runs test cases from testng.xml
 *
 * @author Muaaz Wahid
 * @author Gelin Deng
 * @version 2025-05-11
 */

package week16;
// Log4j libraries
import org.apache.commons.io.FileUtils;
import org.apache.logging.log4j.Level;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.core.config.Configurator;
import org.apache.logging.log4j.core.config.DefaultConfiguration;
// Working with Excel libraries
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
// Selenium libraries
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.WebDriverWait;
// TestNG libraries
import org.testng.Assert;
import org.testng.Reporter;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
// Libraries to working with Excel files
// Imported exceptions and a timer to wait for loading page
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;

/**
 * Web automation script that reads from an Excel file,
 * searches ebay for a camera,
 * and purchases a good deal
 */
public class HolidayShopping {
    private static Logger logger = LogManager.getLogger(HolidayShopping.class);
    private WebDriver driver;
    private Actions actions;
    private double lowestExpectedPrice = 0.00;
    // these 2 fields are used throughout program and do not change
    private final String localPath = "/Projects/cs522/CS522_Selenium/src/week16/";
    private final String screenshotPath = localPath + "screenshots/";

    // Helper method: Remove dollar sign and commas from price
    public static double convertPriceToDouble(String price) {
        String removedCommaStr = price.replace(",", "");
        return Double.parseDouble(removedCommaStr.replace("$", ""));
    }

    // Helper method: Take a screenshot of the website
    public void takeSnapShot(WebDriver webdriver, String fileWithPath) throws Exception {
        // Convert web driver object to TakeScreenshot
        TakesScreenshot scrShot = ((TakesScreenshot)webdriver);
        // Call getScreenshotAs method to create image file
        File sourceFile = scrShot.getScreenshotAs(OutputType.FILE);
        // Move image file to new destination
        File destinationFile = new File(fileWithPath);
        // Copy file at destination
        FileUtils.copyFile(sourceFile, destinationFile);
    }

    /**
     * Setup Log4j, Selenium web driver, navigate to ebay homepage
     */
    @BeforeTest
    void setup() {
        Configurator.initialize(new DefaultConfiguration());
        Configurator.setRootLevel(Level.INFO);
        System.setProperty("webdriver.chrome.driver", "/Projects/cs522/chromedriver");
        driver = new ChromeDriver();
        actions = new Actions(driver);
        String ebayURL = "http://www.ebay.com/";
        driver.get(ebayURL);
        driver.manage().window().maximize();
        String pageTitle = "Electronics, Cars, Fashion, Collectibles & More | eBay";
        Assert.assertEquals(driver.getTitle(), pageTitle);
        logger.info("Step 1: Opened {}", ebayURL);
    }

    /**
     * Reads prices from an Excel sheet,
     * retains the lowest value of the prices for comparison when buying cameras
     * @throws IOException if this method failed to access the Excel file
     */
    @Test(priority = 1)
    void readExcelSheet() throws IOException {
        // Read the prices from Excel file (the prices are the first row)
        FileInputStream fileToRead = new FileInputStream(localPath + "PricingData.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileToRead);
        XSSFSheet sheet = workbook.getSheet("canon");
        int numOfCells = sheet.getRow(0).getLastCellNum();
        // set the lowest expected price to the first cell in the sheet
        lowestExpectedPrice = convertPriceToDouble(sheet.getRow(0).getCell(0).toString());
        Reporter.log("Expected Prices:");
        try {
            // loop through first column of cells
            for (int i = 0; i < numOfCells; i++) {
                String cellStr = sheet.getRow(0).getCell(i).toString();
                if (convertPriceToDouble(cellStr) < lowestExpectedPrice) {
                    lowestExpectedPrice = convertPriceToDouble(cellStr);
                }
                Reporter.log("$" + cellStr);
            }
            // make sure only tenths and hundredths place to right of decimal is preserved
            lowestExpectedPrice = Math.floor(lowestExpectedPrice * 100) / 100;
            Reporter.log("Lowest Expected Price: " + lowestExpectedPrice);
            logger.info("Step 2: Read expected prices from Excel sheet");
        } catch(Throwable t) {
            logger.error("Could not process cell prices from Excel sheet.");
        }
        workbook.close();
        fileToRead.close();
    }

    /**
     * Search for the specific camera on ebay.com
     */
    @Test(priority = 2)
    void searchForProduct() {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(5));
        String mySearchWords = "Canon camera eos 5d mark iv";
        try {
            // screenshot main page with search term
            takeSnapShot(driver, screenshotPath + "test1.png");
            // search for camera in search bar, then hit enter
            Reporter.log("Searching Ebay.com for: " + mySearchWords);
            driver.findElement(By.id("gh-ac")).sendKeys(mySearchWords);
            actions.sendKeys(Keys.ENTER).build().perform();
            wait.until(d -> true);
            Reporter.log("New page:\n" + driver.getTitle());
            // screenshot search results
            takeSnapShot(driver, screenshotPath + "test2.png");
            wait.until(d -> true);
            logger.info("Step 3: Searched for camera, hit enter key, and took a screenshot");
        } catch(Throwable t) {
            logger.error("searchForProduct() failed.");
        }
    }

    /**
     * Searches the first 5 camera results that appear on search
     * Compares to the lowest price from Excel sheet
     * If a camera is lower than the lowest price, it will purchase the camera since it is a good deal
     */
    @Test(priority = 3)
    void purchaseCamera() {
        double currentPrice;
        String priceString = "";
        String xpathString;
        // Search first 5 camera prices results that come up
        for (int i = 1; i <= 5; i++) {
            // use try catch to determine which is the correct xpath of camera price
            try {
                xpathString = "/html/body/div[5]/div[4]/div[3]/div[1]/div[3]/ul/li[" + i +
                        "]/div/div[2]/div[4]/div[1]/div[1]/span";
                priceString = driver.findElement(By.xpath(xpathString)).getText();
            } catch (Throwable t) {
                try {
                    xpathString = "/html/body/div[5]/div[4]/div[3]/div[1]/div[3]/ul/li[" + i +
                            "]/div/div[2]/div[" + 5 + "]/div[1]/div[1]/span";
                    priceString = driver.findElement(By.xpath(xpathString)).getText();
                } catch (Throwable e) {
                    logger.warn("couldn't get camera price");
                }
            }

            Assert.assertNotNull(priceString, "priceString should not be null");
            currentPrice = convertPriceToDouble(priceString);
            Reporter.log("Price of camera " + i + ": " + priceString);
            // compare prices: if lower than lowest expect price, buy it
            if (currentPrice < lowestExpectedPrice) {
                Reporter.log("$" + currentPrice + " is less than lowest expected price $" + lowestExpectedPrice);
                // redirect xpath to selected product link
                xpathString = "/html/body/div[5]/div[4]/div[3]/div[1]/div[3]/ul/li[2]/div/div[2]/a/div/span";
                // try to click product to purchase
                try {
                    driver.findElement(By.xpath(xpathString)).click();
                    // switch to new page and screenshot
                    Object[] windowHandles = driver.getWindowHandles().toArray();
                    driver.switchTo().window((String) windowHandles[1]);
                    takeSnapShot(driver, screenshotPath + "test3.png");

                    // log output and break out of loop since good camera deal found
                    Reporter.log("Successfully bought camera! Navigated to page title:");
                    Reporter.log(driver.getTitle());
                    Reporter.log("Screenshots in dir:");
                    Reporter.log(screenshotPath);
                    logger.info("Step 4: Bought Product");
                    break;
                } catch (Throwable t) {
                    logger.error("Failed to buy product");
                }
            }
        }
    }

    /**
     * Close selenium web driver
     */
    @AfterTest
    void tearDown() {
        driver.quit();
        logger.info("Step 5: Close all open windows and kill all open sessions.");
    }
}
