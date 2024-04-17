import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.List;
import java.util.Locale;

@Component
public class NewsParseTask {

    private final NewsService newsService;

    @Autowired
    public NewsParseTask(NewsService newsService) {
        this.newsService = newsService;
    }

    @Scheduled(fixedDelay = 10000)
    public void parseNews() {
        String url = "https://news.ycombinator.com/";

        try {
            Document doc = Jsoup.connect(url)
                    .userAgent("Mozilla")
                    .timeout(5000)
                    .referrer("https://google.com")
                    .get();
            Elements news = doc.getElementsByClass("storylink");
            for (Element el : news) {
                String title = el.ownText();
                if (!newsService.isExist(title)) {
                    News obj = new News();
                    obj.setTitle(title);
                    obj.setPublicationDate(Calendar.getInstance().getTime());
                    newsService.save(obj);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void generateExcelReport() {
        List<News> allNews = newsService.getAll();
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("News Report");

        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Month");
        headerRow.createCell(1).setCellValue("Number of News");

        Calendar calendar = Calendar.getInstance();
        for (News news : allNews) {
            calendar.setTime(news.getPublicationDate());
            int month = calendar.get(Calendar.MONTH);
            int year = calendar.get(Calendar.YEAR);
            String monthYear = new SimpleDateFormat("MMMM yyyy", Locale.ENGLISH).format(calendar.getTime());

            boolean found = false;
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row.getCell(0).getStringCellValue().equals(monthYear)) {
                    Cell cell = row.getCell(1);
                    cell.setCellValue(cell.getNumericCellValue() + 1);
                    found = true;
                    break;
                }
            }

            if (!found) {
                Row row = sheet.createRow(sheet.getLastRowNum() + 1);
                row.createCell(0).setCellValue(monthYear);
                row.createCell(1).setCellValue(1);
            }
        }

        try (FileOutputStream fileOut = new FileOutputStream("news_report.xlsx")) {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
