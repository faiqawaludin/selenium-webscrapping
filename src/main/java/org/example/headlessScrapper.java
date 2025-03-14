package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class headlessScrapper{
    public static void main(String[] args) {
        System.setProperty("webdriver.chrome.driver", "D:/UNSIKA/FASILKOM/SEM 6/Magenta 2025/Project/Scrapper/src/main/resources/chromedriver/chromedriver.exe");

        ChromeOptions options = new ChromeOptions();
        options.addArguments("--headless");
        options.addArguments("--disable-gpu");
        options.addArguments("--window-size=1920,1080");
        options.addArguments("--ignore-certificate-errors");
        options.addArguments("--disable-extensions");
        options.addArguments("--no-sandbox");
        options.addArguments("--disable-dev-shm-usage");

        WebDriver driver = new ChromeDriver(options);
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        JavascriptExecutor js = (JavascriptExecutor) driver;

        String keyword = "PT Len Industri";
        String searchUrl = "https://www.google.com/search?q=" + keyword + "&tbm=nws";
        List<String[]> newsList = new ArrayList<>();

        try {
            driver.get(searchUrl);
            for (int i = 0; i < 3; i++) {
                if (i > 0) {
                    js.executeScript("window.open()");
                    List<String> tabs = new ArrayList<>(driver.getWindowHandles());
                    driver.switchTo().window(tabs.get(i));
                    driver.get(searchUrl + "&start=" + (i * 10));
                }

                wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[@class='SoaBEf']")));
                List<WebElement> newsElements = driver.findElements(By.xpath("//div[@class='SoaBEf']"));

                for (WebElement news : newsElements) {
                    try {
                        WebElement titleElement = news.findElement(By.xpath(".//div[@role='heading']"));
                        String title = titleElement.getText();
                        WebElement linkElement = news.findElement(By.tagName("a"));
                        String link = linkElement.getAttribute("href");
                        String source = "Tidak ditemukan";
                        try {
                            WebElement sourceElement = news.findElement(By.xpath(".//div[contains(@class, 'MgUUmf')]//span"));
                            source = sourceElement.getText();
                        } catch (NoSuchElementException e) {
                            System.out.println("Sumber berita tidak ditemukan untuk: " + title);
                        }

                        String date = "Tidak ditemukan";
                        try {
                            WebElement dateElement = news.findElement(By.xpath(".//div[contains(@class, 'OSrXXb')]//span"));
                            String rawDate = dateElement.getText();

                            date = convertRelativeDate(rawDate);
                        } catch (NoSuchElementException e) {
                            System.out.println("Tanggal tidak ditemukan untuk: " + title);
                        }

                        newsList.add(new String[]{title, source, date, link});
                    } catch (Exception e) {
                        System.out.println("Gagal mengambil data berita: " + e.getMessage());
                    }
                }
            }
            saveToExcel(newsList, "PT Len News.xlsx");

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            driver.quit();
        }
    }

    private static String convertRelativeDate(String relativeDate) {
        LocalDate today = LocalDate.now();
        LocalDate convertedDate = today;

        Pattern pattern = Pattern.compile("(\\d+)\\s*(hari|minggu|bulan|tahun|jam) lalu");
        Matcher matcher = pattern.matcher(relativeDate.toLowerCase());

        if (matcher.find()) {
            int amount = Integer.parseInt(matcher.group(1));
            String unit = matcher.group(2);

            switch (unit) {
                case "hari":
                    convertedDate = today.minusDays(amount);
                    break;
                case "minggu":
                    convertedDate = today.minusWeeks(amount);
                    break;
                case "bulan":
                    convertedDate = today.minusMonths(amount);
                    break;
                case "tahun":
                    convertedDate = today.minusYears(amount);
                    break;
                case "jam":
                    convertedDate = today;
                    break;
            }
        }

        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("d MMM yyyy");
        return convertedDate.format(formatter);
    }

    public static void saveToExcel(List<String[]> data, String fileName) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("PT Len News");

        String[] headers = {"Judul", "Sumber", "Tanggal", "Link"};
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(getHeaderCellStyle(workbook));
        }

        int rowNum = 1;
        for (String[] rowData : data) {
            Row row = sheet.createRow(rowNum++);
            for (int i = 0; i < rowData.length; i++) {
                row.createCell(i).setCellValue(rowData[i]);
            }
        }

        try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
            workbook.write(fileOut);
            System.out.println("Data berhasil disimpan di " + fileName);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static CellStyle getHeaderCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }
}
