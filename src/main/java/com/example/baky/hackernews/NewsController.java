package com.example.baky.hackernews;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import com.example.baky.hackernews.model.News;
import com.example.baky.hackernews.service.NewsService;

@RestController
public class NewsController {
	
	@Autowired
    private NewsService newsService;

    @GetMapping("/list")
    public List<News> getAllNews() {
        return newsService.getAllNews();
    }

    @PostMapping("/parse")
    public String parseNews() {
        newsService.parseNews();
        return "News parsed successfully!";
    }

    @PostMapping("/report")
    public String generateExcelReport() {
        newsService.generateExcelReport();
        return "Excel report generated successfully!";
    }

    @GetMapping("/excel")
    public ResponseEntity<InputStreamResource> downloadExcelReport() {
        ByteArrayInputStream excelFile = newsService.generateExcelReport();
        HttpHeaders headers = new HttpHeaders();
        headers.add("Content-Disposition", "attachment; filename=news_report.xlsx");

        return ResponseEntity
                .ok()
                .headers(headers)
                .contentType(MediaType.parseMediaType("application/vnd.ms-excel"))
                .body(new InputStreamResource(excelFile));
    }
}
