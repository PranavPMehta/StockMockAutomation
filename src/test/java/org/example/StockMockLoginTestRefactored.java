package org.example;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Logger;

/**
 * Refactored Selenium test for StockMock login and strategy execution workflow.
 * 
 * This test automates the following workflow:
 * 1. Login to StockMock
 * 2. Navigate to basket
 * 3. Select a basket strategy
 * 4. Edit and update strategy
 * 5. Run the strategy
 * 6. Capture overall profit value
 */
public class StockMockLoginTestRefactored {

    private static final Logger LOGGER = Logger.getLogger(StockMockLoginTestRefactored.class.getName());

    // Test Credentials
    private static final String PHONE_NO = "9790395662";
    private static final String PASSWORD = "";

    // URLs
    private static final String BASE_URL = "https://www.stockmock.in";

    // Locators - Phone and Password fields
    private static final By PHONE_INPUT_LOCATOR = By.id("user-phone-no");
    private static final By PASSWORD_INPUT_LOCATOR = By.xpath("//input[@type='password']");
    private static final By LOGIN_BUTTON_LOCATOR = By.xpath("//button[contains(text(), 'LogIn') or contains(text(), 'Login') or contains(text(), 'login')]");

    // Locators - Modal and Navigation
    private static final By CLOSE_BUTTON_LOCATOR = By.xpath("//button[contains(@class, 'close')]");
    private static final By BASKET_BUTTON_LOCATOR = By.xpath("//a[contains(@class, 'header_nav_link') and .//span[contains(text(), 'Basket')]]");

    // Locators - Basket Selection
    private static final String BASKET_ID = "01KDWD18YRS7FRJ5G1Z38VY7WP";
    private static final By BASKET_ITEM_LOCATOR = By.xpath("//li[@data-basket-id='" + BASKET_ID + "']");

    // Locators - Strategy Editing
    private static final By PENCIL_ICON_LOCATOR = By.xpath("//div[@id='basket-strategy-0']//a[@class='fa fa-pencil']");
    private static final By UPDATE_STRATEGY_SAVE_ICON = By.xpath("//button[@class='__button __full__button __run__button']//i[@class='fa fa-save __share__icon']");

    // Locators - Input fields for strategy editing
    private static final By L1_SL_PERCENT_LOCATOR = By.xpath("/html/body/div[1]/div[5]/div[1]/div/div/div[2]/div[3]/div[3]/div[2]/div[2]/div[2]/div[2]/div[1]/input");
    private static final By L2_SL_PERCENT_LOCATOR = By.xpath("/html/body/div[1]/div[5]/div[1]/div/div/div[2]/div[3]/div[4]/div[2]/div[2]/div[2]/div[2]/div[1]/input");
    private static final By ENTRY_TIME_HOUR_LOCATOR = By.xpath("/html/body/div[1]/div[5]/div[1]/div/div/div[2]/div[5]/div[1]/div/div/div/div[3]/div/div[1]/select");
    private static final By ENTRY_TIME_MINUTE_LOCATOR = By.xpath("/html/body/div[1]/div[5]/div[1]/div/div/div[2]/div[5]/div[1]/div/div/div/div[3]/div/div[2]/select");

    // Locators - Settings
    private static final By SETTINGS_DROPDOWN_LOCATOR = By.xpath("/html/body/div[1]/div[4]/div[3]/div[1]/div[4]/div[1]/div/select");

    // Locators - Run Strategy
    private static final By RUN_BUTTON_LOCATOR = By.xpath("//div[@id='basket-strategy-0']//div[@class='strategy_running_status __run']");

    // Locators - Results
    private static final By AVERAGE_CARD_LOCATOR = By.xpath("//div[contains(@class, 'average__card')]");
    private static final By PROFIT_VALUE_LOCATOR = By.xpath(".//div[@class='__value']");

    private WebDriver driver;
    private WebDriverWait wait;
    private static final int DEFAULT_WAIT_TIMEOUT = 10;
    
    // Store test results for Excel export
    private List<TestResult> testResults = new ArrayList<>();
    
    /**
     * Inner class to store test result data
     */
    private static class TestResult {
        int l1SL;
        int l2SL;
        int entryHour;
        int entryMinute;
        String overallProfit;
        String expectancy;
        
        TestResult(int l1SL, int l2SL, int entryHour, int entryMinute, String overallProfit, String expectancy) {
            this.l1SL = l1SL;
            this.l2SL = l2SL;
            this.entryHour = entryHour;
            this.entryMinute = entryMinute;
            this.overallProfit = overallProfit;
            this.expectancy = expectancy;
        }
    }

    @Before
    public void setUp() {
        LOGGER.info("Setting up WebDriver...");
        WebDriverManager.chromedriver().setup();

        ChromeOptions options = new ChromeOptions();
        options.addArguments("start-maximized");
        options.addArguments("disable-blink-features=AutomationControlled");

        this.driver = new ChromeDriver(options);
        this.wait = new WebDriverWait(driver, Duration.ofSeconds(DEFAULT_WAIT_TIMEOUT));

        LOGGER.info("WebDriver initialized successfully");
    }

    @Test
    public void testStockMockLoginAndStrategyExecution() {
        try {
            navigateToStockMock();
            login();
            closeLoginModal();
            navigateToBasket();
            selectBasketStrategy();
            configureWeekdaySetting();

            // Generate entry times from 9:16 to 12:00 (incrementing by 1 minute)
            int[][] entryTimes = generateTimeRange(9, 16, 12, 0);
            
            // Loop through SL % values (5 to 100) and entry times
            for (int slPercent = 5; slPercent <= 100; slPercent++) {
                for (int[] entryTime : entryTimes) {
                    int hour = entryTime[0];
                    int minute = entryTime[1];

                    LOGGER.info("========================================");
                    LOGGER.info("Testing with SL%: " + slPercent + ", Entry Time: " + hour + ":" + (minute < 10 ? "0" : "") + minute);
                    LOGGER.info("========================================");

                    editAndUpdateStrategy(slPercent, hour, minute);
                    runStrategy();
                    captureResults(slPercent, hour, minute);

                    LOGGER.info("Completed iteration: SL%=" + slPercent + ", Entry Time=" + hour + ":" + (minute < 10 ? "0" : "") + minute);
                    LOGGER.info("");
                }
            }

            // Export results to Excel
            exportResultsToExcel();

            LOGGER.info("Test completed successfully!");
        } catch (Exception e) {
            LOGGER.severe("Test failed: " + e.getMessage());
            e.printStackTrace();
        }
    }

    /**
     * Navigate to StockMock homepage
     */
    private void navigateToStockMock() {
        LOGGER.info("Navigating to " + BASE_URL);
        driver.navigate().to(BASE_URL);
        sleep(2000); // Allow page to load
        LOGGER.info("Successfully navigated to StockMock");
    }

    /**
     * Login to StockMock with provided credentials
     */
    private void login() {
        LOGGER.info("Attempting login with phone: " + PHONE_NO);

        WebElement phoneInput = wait.until(ExpectedConditions.presenceOfElementLocated(PHONE_INPUT_LOCATOR));
        phoneInput.clear();
        phoneInput.sendKeys(PHONE_NO);
        LOGGER.info("Entered phone number");

        WebElement passwordInput = wait.until(ExpectedConditions.presenceOfElementLocated(PASSWORD_INPUT_LOCATOR));
        passwordInput.clear();
        passwordInput.sendKeys(PASSWORD);
        LOGGER.info("Entered password");

        WebElement loginButton = wait.until(ExpectedConditions.elementToBeClickable(LOGIN_BUTTON_LOCATOR));
        loginButton.click();
        LOGGER.info("Clicked login button");

        sleep(100); // Wait for login to complete
        LOGGER.info("Login successful");
    }

    /**
     * Close the modal dialog that appears after login
     */
    private void closeLoginModal() {
        LOGGER.info("Attempting to close login modal...");
        try {
            WebElement closeButton = wait.until(ExpectedConditions.elementToBeClickable(CLOSE_BUTTON_LOCATOR));
            closeButton.click();
            LOGGER.info("Login modal closed successfully");
        } catch (Exception e) {
            LOGGER.warning("Could not close modal or modal not present: " + e.getMessage());
        }
    }

    /**
     * Navigate to the basket page via navbar button
     */
    private void navigateToBasket() {
        LOGGER.info("Navigating to basket...");
        WebElement basketButton = wait.until(ExpectedConditions.elementToBeClickable(BASKET_BUTTON_LOCATOR));
        basketButton.click();
        LOGGER.info("Successfully navigated to basket");
    }

    /**
     * Select the specific basket strategy by ID
     */
    private void selectBasketStrategy() {
        LOGGER.info("Selecting basket with ID: " + BASKET_ID);
        WebElement basketItem = wait.until(ExpectedConditions.elementToBeClickable(BASKET_ITEM_LOCATOR));
        basketItem.click();
        sleep(500);
        LOGGER.info("Basket strategy selected");
    }

    /**
     * Configure the settings dropdown to select "Weekday" option
     */
    private void configureWeekdaySetting() {
        LOGGER.info("Configuring Weekday setting...");
        try {
            WebElement settingsDropdown = wait.until(ExpectedConditions.presenceOfElementLocated(SETTINGS_DROPDOWN_LOCATOR));
            
            // Scroll to element to ensure it's visible
            ((org.openqa.selenium.JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", settingsDropdown);
            sleep(100);
            
            // Click to open dropdown
            settingsDropdown.click();
            sleep(500);
            
            // Find and select the Weekday option
            try {
                WebElement weekdayOption = settingsDropdown.findElement(By.xpath(".//option[contains(text(), 'Weekday')]"));
                weekdayOption.click();
                LOGGER.info("Weekday option selected");
            } catch (Exception e) {
                // Try alternative selector
                WebElement weekdayOption = settingsDropdown.findElement(By.xpath(".//option[@value='Weekday']"));
                weekdayOption.click();
                LOGGER.info("Weekday option selected (by value)");
            }
            
            sleep(500);
            LOGGER.info("Weekday setting configured successfully");
        } catch (Exception e) {
            LOGGER.severe("Error configuring Weekday setting: " + e.getMessage());
            e.printStackTrace();
        }
    }

    /**
     * Generate an array of time ranges from start time to end time, incrementing by 1 minute
     * @param startHour starting hour (e.g., 9)
     * @param startMinute starting minute (e.g., 16)
     * @param endHour ending hour (e.g., 12)
     * @param endMinute ending minute (e.g., 0)
     * @return 2D array of [hour, minute] pairs
     */
    private int[][] generateTimeRange(int startHour, int startMinute, int endHour, int endMinute) {
        LOGGER.info("Generating time range from " + startHour + ":" + (startMinute < 10 ? "0" : "") + startMinute + 
                   " to " + endHour + ":" + (endMinute < 10 ? "0" : "") + endMinute);
        
        // Calculate total minutes from start to end
        int startTotalMinutes = startHour * 60 + startMinute;
        int endTotalMinutes = endHour * 60 + endMinute;
        
        // If end time is on same day or next day, adjust calculation
        if (endTotalMinutes <= startTotalMinutes) {
            endTotalMinutes += 24 * 60; // Add 24 hours for next day
        }
        
        int totalSlots = endTotalMinutes - startTotalMinutes + 1;
        int[][] times = new int[totalSlots][2];
        
        int currentTotalMinutes = startTotalMinutes;
        for (int i = 0; i < totalSlots; i++) {
            times[i][0] = (currentTotalMinutes / 60) % 24; // Hour
            times[i][1] = currentTotalMinutes % 60; // Minute
            currentTotalMinutes++;
        }
        
        LOGGER.info("Generated " + totalSlots + " time slots");
        return times;
    }

    /**
     * Edit and update the selected strategy, including changing SL % and entry time
     * @param slPercent the SL % value to set for both L1 and L2
     * @param hour the hour for entry time
     * @param minute the minute for entry time
     */
    private void editAndUpdateStrategy(int slPercent, int hour, int minute) {
        LOGGER.info("Editing strategy...");
        // Click pencil icon to edit
        WebElement pencilIcon = wait.until(ExpectedConditions.elementToBeClickable(PENCIL_ICON_LOCATOR));
        pencilIcon.click();
        sleep(100); // Wait longer for edit form to load
        LOGGER.info("Pencil icon clicked");

        // Change SL % for L1
        changeSLPercentForLeg("L1", slPercent);
        // Change SL % for L2
        changeSLPercentForLeg("L2", slPercent);
        // Change entry time
        changeEntryTime(hour, minute);

        // Click update strategy save icon
        WebElement saveIcon = wait.until(ExpectedConditions.elementToBeClickable(UPDATE_STRATEGY_SAVE_ICON));
        saveIcon.click();
        sleep(100);
        LOGGER.info("Strategy updated");

        // Check for confirmation modal and handle if present
        handleConfirmationModalIfPresent();
    }

    /**
     * Change SL % for a given leg (L1 or L2)
     */
    private void changeSLPercentForLeg(String legName, int slPercent) {
        LOGGER.info("Changing SL % for " + legName + " to " + slPercent);
        try {
            By locator;
            if ("L1".equals(legName)) {
                locator = L1_SL_PERCENT_LOCATOR;
            } else if ("L2".equals(legName)) {
                locator = L2_SL_PERCENT_LOCATOR;
            } else {
                LOGGER.warning("Unknown leg name: " + legName);
                return;
            }
            
            WebElement slInput = wait.until(ExpectedConditions.presenceOfElementLocated(locator));
            
            // Scroll to element to ensure it's visible
            ((org.openqa.selenium.JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", slInput);
            sleep(100);
            
            // Click to focus on the element
            slInput.click();
            sleep(100);
            
            // Clear the field using multiple approaches to ensure complete clearing
            // 1. Select all text
            slInput.sendKeys(org.openqa.selenium.Keys.chord(org.openqa.selenium.Keys.CONTROL, "a"));
            sleep(50);
            
            // 2. Delete selected text
            slInput.sendKeys(org.openqa.selenium.Keys.DELETE);
            sleep(50);
            
            // 3. Also clear using JavaScript as backup
            ((org.openqa.selenium.JavascriptExecutor) driver).executeScript("arguments[0].value = '';", slInput);
            sleep(50);
            
            // 4. Trigger input event to notify form of change
            ((org.openqa.selenium.JavascriptExecutor) driver).executeScript(
                "arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", slInput);
            sleep(100);
            
            // 5. Verify the field is actually empty
            String currentValue = (String) ((org.openqa.selenium.JavascriptExecutor) driver)
                .executeScript("return arguments[0].value;", slInput);
            LOGGER.info("Field value after clearing: '" + currentValue + "'");
            
            // 6. Type the new value
            slInput.sendKeys(String.valueOf(slPercent));
            sleep(100);
            
            // 7. Trigger change event to notify form
            ((org.openqa.selenium.JavascriptExecutor) driver).executeScript(
                "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", slInput);
            sleep(100);
            
            // 8. Verify the new value was set
            String newValue = (String) ((org.openqa.selenium.JavascriptExecutor) driver)
                .executeScript("return arguments[0].value;", slInput);
            LOGGER.info("Field value after setting: '" + newValue + "'");
            
            LOGGER.info("Successfully changed SL % for " + legName + " to " + slPercent);
        } catch (Exception e) {
            LOGGER.warning("Error changing SL % for " + legName + ": " + e.getMessage());
            e.printStackTrace();
        }
    }

    /**
     * Change entry time to given hour and minute
     */
    private void changeEntryTime(int hour, int minute) {
        LOGGER.info("Changing entry time to " + hour + ":" + minute);
        try {
            sleep(100); // Wait for form elements to be interactive
            
            // Change hour
            WebElement hourSelect = wait.until(ExpectedConditions.presenceOfElementLocated(ENTRY_TIME_HOUR_LOCATOR));
            
            // Scroll to element to ensure it's visible
            ((org.openqa.selenium.JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", hourSelect);
            sleep(100);
            
            // Click to open dropdown
            hourSelect.click();
            sleep(100);
            
            // Find and select the hour option
            try {
                WebElement hourOption = hourSelect.findElement(By.xpath(".//option[@value='" + hour + "']"));
                hourOption.click();
            } catch (Exception e) {
                // Try by text if value attribute doesn't match
                WebElement hourOption = hourSelect.findElement(By.xpath(".//option[contains(text(), '" + hour + "')]"));
                hourOption.click();
            }
            LOGGER.info("Hour set to " + hour);
            
            // Change minute
            sleep(100);
            WebElement minuteSelect = wait.until(ExpectedConditions.presenceOfElementLocated(ENTRY_TIME_MINUTE_LOCATOR));
            
            // Scroll to element to ensure it's visible
            ((org.openqa.selenium.JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", minuteSelect);
            sleep(100);
            
            // Click to open dropdown
            minuteSelect.click();
            sleep(100);
            
            // Find and select the minute option
            try {
                WebElement minuteOption = minuteSelect.findElement(By.xpath(".//option[@value='" + minute + "']"));
                minuteOption.click();
            } catch (Exception e) {
                // Try by text if value attribute doesn't match
                WebElement minuteOption = minuteSelect.findElement(By.xpath(".//option[contains(text(), '" + minute + "')]"));
                minuteOption.click();
            }
            LOGGER.info("Minute set to " + minute);
            
            LOGGER.info("Entry time changed to " + hour + ":" + minute);
        } catch (Exception e) {
            LOGGER.warning("Error changing entry time: " + e.getMessage());
            e.printStackTrace();
        }
    }

    /**
     * Handle the confirmation modal that may appear after update
     */
    private void handleConfirmationModalIfPresent() {
        LOGGER.info("Checking for confirmation modal...");
        try {
            // Use exact XPath for the Update button
            By modalUpdateButtonLocator = By.xpath("/html/body/div[1]/div[6]/div[1]/div/div/div[2]/button[2]");
            WebElement updateModalButton = wait.until(ExpectedConditions.elementToBeClickable(modalUpdateButtonLocator));
            updateModalButton.click();
            sleep(500);
            LOGGER.info("Confirmation modal handled");
        } catch (Exception e) {
            LOGGER.info("No confirmation modal present: " + e.getMessage());
        }
    }

    /**
     * Run the strategy by clicking the run button
     */
    private void runStrategy() {
        LOGGER.info("Running strategy...");
        sleep(500); // Wait for page to stabilize after update
        
        WebElement runButton = wait.until(ExpectedConditions.elementToBeClickable(RUN_BUTTON_LOCATOR));
        runButton.click();
        sleep(4000); // Wait for strategy execution
        LOGGER.info("Strategy executed");
    }

    /**
     * Capture both overall profit and expectancy values from the results
     */
    private void captureResults(int slPercent, int hour, int minute) {
        LOGGER.info("Capturing overall profit and expectancy values...");
        sleep(1000); // Wait for results to render

        try {
            String overallProfit = captureOverallProfitValue();
            String expectancy = captureExpectancyValue();
            LOGGER.info("Capturing overall profit and expectancy values...");

            // Store the result
            testResults.add(new TestResult(slPercent, slPercent, hour, minute, overallProfit, expectancy));

            logCapturedResults(slPercent, hour, minute, overallProfit, expectancy);
        } catch (Exception e) {
            LOGGER.severe("Error capturing results: " + e.getMessage());
            e.printStackTrace();
        }
    }

    /**
     * Capture and return the overall profit value from the results
     */
    private String captureOverallProfitValue() {
        LOGGER.info("Capturing overall profit value...");

        try {
            WebElement overallProfitCard = findOverallProfitCard();
            WebElement profitValueElement = overallProfitCard.findElement(PROFIT_VALUE_LOCATOR);
            String overallProfit = profitValueElement.getText();
            LOGGER.info("Overall Profit Value: " + overallProfit);
            return overallProfit;
        } catch (Exception e) {
            LOGGER.severe("Could not capture profit value: " + e.getMessage());
            return "N/A";
        }
    }

    /**
     * Capture and return the expectancy value from the results
     */
    private String captureExpectancyValue() {
        LOGGER.info("Capturing expectancy value...");

        try {
            // Find the Expectancy card
            WebElement expectancyCard = findExpectancyCard();
            WebElement expectancyValueElement = expectancyCard.findElement(PROFIT_VALUE_LOCATOR);
            String expectancy = expectancyValueElement.getText();
            LOGGER.info("Expectancy Value: " + expectancy);
            return expectancy;
        } catch (Exception e) {
            LOGGER.severe("Could not capture expectancy value: " + e.getMessage());
            return "N/A";
        }
    }

    /**
     * Find the average__card div containing "Overall profit"
     * @return WebElement representing the overall profit card
     */
    private WebElement findOverallProfitCard() {
        LOGGER.info("Finding overall profit card...");
        
        try {
            // Try with exact text match
            return driver.findElement(By.xpath("//div[contains(@class, 'average__card') and .//div[@class='__title' and contains(text(), 'Overall profit')]]"));
        } catch (Exception e1) {
            try {
                // Try with case-insensitive match
                return driver.findElement(By.xpath("//div[contains(@class, 'average__card') and .//div[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'overall profit')]]"));
            } catch (Exception e2) {
                // Fallback: iterate through cards
                return findCardByIteration("Overall profit");
            }
        }
    }

    /**
     * Find the average__card div containing "Expectancy"
     * @return WebElement representing the expectancy card
     */
    private WebElement findExpectancyCard() {
        LOGGER.info("Finding expectancy card...");
        
        try {
            // Try with exact text match
            return driver.findElement(By.xpath("//div[contains(@class, 'average__card') and .//div[@class='__title' and contains(text(), 'Expectancy')]]"));
        } catch (Exception e1) {
            try {
                // Try with case-insensitive match
                return driver.findElement(By.xpath("//div[contains(@class, 'average__card') and .//div[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'expectancy')]]"));
            } catch (Exception e2) {
                // Fallback: iterate through cards
                return findCardByIteration("Expectancy");
            }
        }
    }

    /**
     * Find average__card by iterating through all elements
     * @return WebElement representing the card
     */
    private WebElement findCardByIteration(String cardTitle) {
        LOGGER.info("Falling back to card iteration for: " + cardTitle);
        
        List<WebElement> cards = driver.findElements(AVERAGE_CARD_LOCATOR);
        LOGGER.info("Found " + cards.size() + " average__card elements");

        for (WebElement card : cards) {
            try {
                WebElement title = card.findElement(By.xpath(".//div[@class='__title']"));
                if (title.getText().contains(cardTitle)) {
                    LOGGER.info(cardTitle + " card found by iteration");
                    return card;
                }
            } catch (Exception e) {
                // Continue searching
            }
        }

        throw new RuntimeException(cardTitle + " card not found");
    }

    /**
     * Sleep for specified milliseconds
     * @param millis milliseconds to sleep
     */
    private void sleep(long millis) {
        try {
            Thread.sleep(millis);
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            LOGGER.warning("Sleep interrupted: " + e.getMessage());
        }
    }

    /**
     * Log the captured results in a formatted manner
     */
    private void logCapturedResults(int slPercent, int hour, int minute, String overallProfit, String expectancy) {
        String separator = "========================================";
        LOGGER.info(separator);
        LOGGER.info("TEST RESULT SUMMARY");
        LOGGER.info("L1 SL%: " + slPercent);
        LOGGER.info("L2 SL%: " + slPercent);
        LOGGER.info("Entry Time: " + hour + ":" + (minute < 10 ? "0" : "") + minute);
        LOGGER.info("Overall Profit: " + overallProfit);
        LOGGER.info("Expectancy: " + expectancy);
        LOGGER.info(separator);
    }

    /**
     * Export test results to Excel file
     */
    private void exportResultsToExcel() {
        LOGGER.info("Exporting results to Excel...");
        
        try {
            org.apache.poi.xssf.usermodel.XSSFWorkbook workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook();
            org.apache.poi.ss.usermodel.Sheet sheet = workbook.createSheet("Strategy Results");
            
            // Create header row
            org.apache.poi.ss.usermodel.Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("L1 SL%");
            headerRow.createCell(1).setCellValue("L2 SL%");
            headerRow.createCell(2).setCellValue("Entry Time");
            headerRow.createCell(3).setCellValue("Overall Profit");
            headerRow.createCell(4).setCellValue("Expectancy");
            
            // Add data rows
            int rowNum = 1;
            for (TestResult result : testResults) {
                org.apache.poi.ss.usermodel.Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(result.l1SL);
                row.createCell(1).setCellValue(result.l2SL);
                // Combine hour and minute in HH:MM format
                String entryTime = String.format("%02d:%02d", result.entryHour, result.entryMinute);
                row.createCell(2).setCellValue(entryTime);
                row.createCell(3).setCellValue(result.overallProfit);
                row.createCell(4).setCellValue(result.expectancy);
            }
            
            // Adjust column widths
            sheet.autoSizeColumn(0);
            sheet.autoSizeColumn(1);
            sheet.autoSizeColumn(2);
            sheet.autoSizeColumn(3);
            sheet.autoSizeColumn(4);
            
            // Write to file
            String filePath = "target/StrategyTestResults.xlsx";
            java.io.FileOutputStream fileOutputStream = new java.io.FileOutputStream(filePath);
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            workbook.close();
            
            LOGGER.info("Results exported successfully to: " + filePath);
        } catch (java.io.IOException e) {
            LOGGER.severe("Error exporting results to Excel: " + e.getMessage());
            e.printStackTrace();
        }
    }

    @After
    public void tearDown() {
        LOGGER.info("Test completed. Browser remains open for inspection.");
        // Uncomment below to close browser automatically
        // if (driver != null) {
        //     driver.quit();
        //     LOGGER.info("Browser closed");
        // }
    }
}
