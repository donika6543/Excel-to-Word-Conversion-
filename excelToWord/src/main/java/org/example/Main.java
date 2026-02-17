package org.example;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;

public class Main {

    public static void main(String[] args) {

        String excelFilePath = "src/data/Employee details.xlsx";
        String wordFilePath = "src/data/Employee list.docx";

        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis);
             XWPFDocument document = new XWPFDocument()) {

            Sheet sheet = workbook.getSheetAt(0);

            int numberOfColumns = sheet.getRow(0).getLastCellNum();
            int numberOfRows = sheet.getPhysicalNumberOfRows();


            XWPFTable table = document.createTable(numberOfRows, numberOfColumns);

            int rowIndex = 0;

            for (Row excelRow : sheet) {

                XWPFTableRow wordRow = table.getRow(rowIndex);

                for (int colIndex = 0; colIndex < numberOfColumns; colIndex++) {

                    Cell cell = excelRow.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    String cellValue = getCellValue(cell);

                    wordRow.getCell(colIndex).setText(cellValue);
                }

                rowIndex++;
            }

            try (FileOutputStream fos = new FileOutputStream(wordFilePath)) {
                document.write(fos);
            }

            System.out.println("Excel data successfully converted to Word TABLE!");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static String getCellValue(Cell cell) {

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                }
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
}


