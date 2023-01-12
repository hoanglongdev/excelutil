package com.kupin.service;

import com.kupin.excelutil.BaseExcelExporter;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.util.List;
import java.util.Random;

@Service
public class ExcelExporterService extends BaseExcelExporter<Integer> {
    private static final Log logger = LogFactory.getLog(ExcelExporterService.class);
    private static final String NAME = "abcdefghijklmnopqrstuvwxyz";
    private static final Random RANDOM = new Random();
    private int rowCount;

    public ExcelExporterService() {
        super("template.xlsx");
    }

    public void setRowCount(int rowCount) {
        this.rowCount = rowCount;
    }

    @Override
    protected void doExport(List<Integer> objectList) throws IOException {
        for (int i = 0; i < rowCount; i++) {
            writeObject(currentRow);
            zipWriter.flush();
            if (currentRow % 10000 == 0) {
                logger.info("Written " + i + " rows after " + ((System.currentTimeMillis() - startTime) / 1000) + " seconds");
            }
        }
        logger.info("Written " + rowCount + " rows after " + ((System.currentTimeMillis() - startTime) / 1000) + " seconds");
    }

    @Override
    protected void writeObject(Integer object) throws IOException {
        createRow();
        createCell(0, currentRow);
        createCell(1, randomName());
        createCell(2, randomPhone());
        createCell(3, randomEmail());
        endRow();
    }

    private String randomName() {
        return randomString(true) + " " + randomString(true);
    }

    private String randomString(boolean upperFirst) {
        StringBuilder sb = new StringBuilder();
        int length = RANDOM.nextInt(5) + 3;
        for (int i = 0; i < length; i++) {
            if (i == 0 && upperFirst) {
                sb.append(Character.toUpperCase(NAME.charAt(RANDOM.nextInt(NAME.length()))));
            } else {
                sb.append(NAME.charAt(RANDOM.nextInt(NAME.length())));
            }
        }
        return sb.toString();
    }

    private String randomPhone() {
        StringBuilder sb = new StringBuilder();
        sb.append("0");
        for (int i = 0; i < 8; i++) {
            sb.append(RANDOM.nextInt(9));
        }
        return sb.toString();
    }

    private String randomEmail() {
        return randomString(false) + "@kupin.com";
    }
}
