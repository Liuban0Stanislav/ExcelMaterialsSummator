package org.lyuban.apachepoi.unclassified;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
public class ExcelBookCreator {
    private Workbook workbook;
    private Sheet sheet;
    public ExcelBookCreator() {
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet();
    }

    public ExcelBookCreator(String file, int numOfSheet) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        workbook = new XSSFWorkbook(fis);
        sheet = workbook.getSheetAt(numOfSheet);
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    public Sheet getSheet() {
        return sheet;
    }

    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }
}
