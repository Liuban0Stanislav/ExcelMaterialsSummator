package org.lyuban.apachepoi.servises;


import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.lyuban.apachepoi.unclassified.ExcelBookCreator;
import org.lyuban.apachepoi.constants.Constants;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.Collection;
import java.util.Date;
import java.util.HashSet;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * Сервис извлекает данные из книги.
 */
@Service
@Slf4j
public class ExcelMaterialsExtractorService {
    private static Collection<String> materials = new HashSet<>();

    public static void main(String[] args) throws IOException {
//        writeIntoExcel("D:/testExcel.xls");
//        getNamesFromExcel("D:/Работа/ИЗДЕЛИЯ/02 ТЕХНИЧЕСКИЕ ПРЕДЛОЖЕНИЯ/КБГР.0027 после командировки/КБГР.0027В СП .xlsx");
//        getNamesFromExcel("D:/testExcel.xls");
        getAllMaterials("D:/testMaterials.xlsx", 0);
//        getMaterialsCollection("D:/testMaterials.xlsx", 0, 3, 1);
        printMaterialsCollection();
    }

    public static void printMaterialsCollection() {
        //Класс  AtomicInteger  представляет собой целочисленное значение, которое можно изменять атомарно.
        //Это означает, что операции инкремента, декремента и обновления значения будут происходить
        // без возможности конфликтов при многопоточном доступе.
        AtomicInteger counter = new AtomicInteger(1);
        materials.stream()
                .sorted()
                .forEach(el -> System.out.println((counter.getAndIncrement()) + " " + el));
    }

    public static void writeIntoExcel(String file) throws FileNotFoundException, IOException {
        Workbook book = new HSSFWorkbook();
        Sheet sheet = book.createSheet("Birthdays");

        // Нумерация начинается с нуля
        Row row = sheet.createRow(0);

        // Мы запишем имя и дату в два столбца
        // имя будет String, а дата рождения --- Date,
        // формата dd.mm.yyyy
        Cell name = row.createCell(0); //Cell то ячейка
        name.setCellValue("John");

        Cell birthdate = row.createCell(1);
        DataFormat format = book.createDataFormat();
        CellStyle dateStyle = book.createCellStyle();
        dateStyle.setDataFormat(format.getFormat("dd.mm.yyyy"));
        birthdate.setCellStyle(dateStyle);


        // Нумерация лет начинается с 1900-го
        birthdate.setCellValue(new Date(110, 10, 10));

        // Меняем размер столбца
        sheet.autoSizeColumn(1);

        // Записываем всё в файл
        book.write(new FileOutputStream(file));
        book.close();
    }

    public static void getNamesFromExcel(String file) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        Workbook wb = new HSSFWorkbook(fis);
//        Workbook wb = new XSSFWorkbook(new FileInputStream(file));
        Sheet sheet = wb.getSheetAt(0);

        for (Row row : sheet) {
            for (Cell cell : row) {
                CellType cellType = cell.getCellType();
                if (cellType == CellType.STRING) {
                    System.out.println(cell.getStringCellValue().toString());

                    int currentRowIndex = cell.getRowIndex();
                    int currentColumnIndex = cell.getColumnIndex();
                    if (cell.getStringCellValue().equals("John")) {
                        Cell newCell = row.createCell(currentColumnIndex + 2);//Создаем новую ячейку
                        newCell.setCellValue("Новое значение2");//записываем туда навое занчение
                    }
                }
                if (cellType == CellType.NUMERIC) {
                    System.out.println(cell.getNumericCellValue());
                }
            }
        }
        fis.close();

        FileOutputStream fos = new FileOutputStream(file);
        wb.write(fos);
        fos.close();
        wb.close();
    }

    /**
     * Метод находит все метериалы, которые использовались в проекте и заполняет или set
     * {@link ExcelMaterialsExtractorService#materials}.
     * <br>Проходим по строкам сверху в низ и по ячейкам слева на право.
     * Если находим в ячейке строковую запись {@value Constants#NAIMENOVANIE},
     * то переходим на следующую строку в ячейку с такми же номером и берем из нее значение и кладем в set
     * {@link ExcelMaterialsExtractorService#materials}.</br>
     * Продоложаем повторять это до тех пор пока не встретим пустую ячейку.
     *
     * @param file fадрес читаемой книги
     * @throws IOException
     */
    public static void getAllMaterials(String file, int numOfSheet) throws IOException {
        Sheet sheet = getSheet(file, numOfSheet);
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().equals(Constants.NAIMENOVANIE)) {
                    int nextRowNum = row.getRowNum() + 1;
//                    System.out.println("nextRowNum = " + nextRowNum);

                    int currentColumn = cell.getColumnIndex();
//                    System.out.println("currentColumn = " + currentColumn);

                    getMaterialsCollection(file, numOfSheet, nextRowNum, currentColumn);
                }
            }
        }
    }

    /**
     * Метод проходит по строкам сверху вниз, и заполняет коллекцию строками с материалами, до тех пор пока,
     * не встретит пустую строку.
     *
     * @param file          адрес читаемой книги.
     * @param numOfSheet    номер нужного листа.
     * @param nextRowNum    номер следующей строки, откуда будет читаться значение материала.
     * @param currentColumn номер колонки в которой находится материал.
     * @throws IOException
     */
    private static void getMaterialsCollection(String file, int numOfSheet, int nextRowNum, int currentColumn) throws IOException {
        Sheet sheet = getSheet(file, numOfSheet);
        while (!isRowEmpty(nextRowNum, file, numOfSheet)) {
            Row row = sheet.getRow(nextRowNum);
            String value = row.getCell(currentColumn) != null ? row.getCell(currentColumn).getStringCellValue() : null;
            materials.add(value);
            nextRowNum++;
        }
    }

    /**
     * Метод проверяет является ли строка пустой. Если строка пуста то возвращает true? Иначе false.
     *
     * @param rowNum     номер строки.
     * @param file       адрес файла
     * @param numOfSheet номер листа
     * @return true или false
     */
    private static boolean isRowEmpty(int rowNum, String file, int numOfSheet) throws IOException {
        Sheet sheet = getSheet(file, numOfSheet);
        if (sheet.getRow(rowNum) == null) {
            return true;
        }
        return false;
    }

    /**
     * Метод создает новую книгу и новый лист в ней. Если не задать имя листа, то присвоится занчение по умолчанию:
     * "Лист1".
     *
     * @param file        адрес нахождения новой книги.
     * @param nameOfSheet имя нового листа.
     * @return новый лист
     * @throws IOException
     */
    private static Sheet createSheet(String file, String nameOfSheet) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        Workbook wb = new HSSFWorkbook(fis);
        if (nameOfSheet == null) {
            return wb.createSheet("Лист1");
        }
        return wb.createSheet(nameOfSheet);
    }

    /**
     * Метод получает существующую книгу по указанному адресу и возвращает ее заданный лист по номеру.
     *
     * @param file       место нахождения книги.
     * @param numOfSheet номер листа, который надо вернуть.
     * @return лист книги
     * @throws IOException
     */
    private static Sheet getSheet(String file, int numOfSheet) throws IOException {
        ExcelBookCreator excelBookCreator = new ExcelBookCreator(file, numOfSheet);
        return excelBookCreator.getSheet();
    }
}
