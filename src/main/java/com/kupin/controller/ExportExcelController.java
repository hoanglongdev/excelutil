package com.kupin.controller;

import com.kupin.service.ExcelExporterService;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

@Controller
public class ExportExcelController {

    private final ExcelExporterService testExcelExporter;

    public ExportExcelController(ExcelExporterService testExcelExporter) {
        this.testExcelExporter = testExcelExporter;
    }

    @RequestMapping("/")
    public String welcome() {
        return "welcome";
    }

    @PostMapping("/export")
    public void export(HttpServletResponse response, @RequestParam int rowCount) throws IOException {
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=\"one-million-row.xlsx\"");
        testExcelExporter.setRowCount(rowCount);
        testExcelExporter.export(response.getOutputStream(), null);
    }
}
