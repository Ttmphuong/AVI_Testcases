package org.example;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;

public class Main {
    public static void main(String[] args) throws IOException {
        String projectPath = System.getProperty("user.dir");
        String importFilePath = projectPath + File.separator + "importFile" + File.separator + "Mersive Solstice Pod.xlsx";
        FileInputStream file = new FileInputStream(importFilePath);
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);

        try {
            writeTestExpectedResult(sheet);

            FileOutputStream output = new FileOutputStream(importFilePath);
            workbook.write(output);

            // Close resources
            output.close();
            workbook.close();
            file.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String getStringCellValue(Cell cell) {
        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue();
        } else if (cell.getCellType() == CellType.NUMERIC) {
            return String.valueOf(cell.getNumericCellValue());
        } else if (cell.getCellType() == CellType.BOOLEAN) {
            return String.valueOf(cell.getBooleanCellValue());
        } else {
            return "";
        }
    }

    private static void writeTcTitle(Sheet sheet) throws IOException {


        int sourceColumnIndex = 2; // Index of the source column (zero-based)
        int destinationColumnIndex = 4; // Index of the destination column (zero-based)

        // Read data from rows 4 to 60 in the source column and write it to the destination column
        int startRow = 2; // Row index to start reading (zero-based)
        int endRow = 59; // Row index to end reading (zero-based)
        for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
            Row sourceRow = sheet.getRow(rowIndex);
            if (sourceRow != null) {
                Cell sourceCell = sourceRow.getCell(sourceColumnIndex);
                if (sourceCell != null) {
                    // Retrieve the value from the source cell
                    String cellValue = getStringCellValue(sourceCell);

                    // Write the value to the destination column
                    Row destinationRow = sheet.getRow(rowIndex);
                    if (destinationRow == null) {
                        destinationRow = sheet.createRow(rowIndex);
                    }
                    Cell destinationCell = destinationRow.createCell(destinationColumnIndex);
                    destinationCell.setCellValue("Verify " + cellValue + " value while looking at device Management portal");
                }
            }
        }
    }
    private static void writeTestStep(Sheet sheet) throws IOException {


        int sourceColumnIndex = 2; // Index of the source column (zero-based)
        int destinationColumnIndex = 6; // Index of the destination column (zero-based)

        // Read data from rows 4 to 60 in the source column and write it to the destination column
        int startRow = 2; // Row index to start reading (zero-based)
        int endRow = 59; // Row index to end reading (zero-based)
        for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
            Row sourceRow = sheet.getRow(rowIndex);
            if (sourceRow != null) {
                Cell sourceCell = sourceRow.getCell(sourceColumnIndex);
                if (sourceCell != null) {
                    // Retrieve the value from the source cell
                    String cellValue = getStringCellValue(sourceCell);

                    // Write the value to the destination column
                    Row destinationRow = sheet.getRow(rowIndex);
                    if (destinationRow == null) {
                        destinationRow = sheet.createRow(rowIndex);
                    }
                    Cell destinationCell = destinationRow.createCell(destinationColumnIndex);
                    destinationCell.setCellValue( "1. Login Symphony \n2.Go to Mersive Solstice Pod device under test \n3. Go to Extended properties tab of the device \n4. Check the " + cellValue + " value \n5. Open web UI of Mersive Solstice Pod on another browser \n6. Go to Licensing  tab \n7.Verify that the " + cellValue + " show on Live monitoring tab will be correct");
                }
            }
        }
    }
    private static void writeTestExpectedResult(Sheet sheet) throws IOException {


        int sourceColumnIndex = 2; // Index of the source column (zero-based)
        int destinationColumnIndex = 8; // Index of the destination column (zero-based)

        // Read data from rows 4 to 60 in the source column and write it to the destination column
        int startRow = 2; // Row index to start reading (zero-based)
        int endRow = 59; // Row index to end reading (zero-based)
        for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
            Row sourceRow = sheet.getRow(rowIndex);
            if (sourceRow != null) {
                Cell sourceCell = sourceRow.getCell(sourceColumnIndex);
                if (sourceCell != null) {
                    // Retrieve the value from the source cell
                    String cellValue = getStringCellValue(sourceCell);

                    // Write the value to the destination column
                    Row destinationRow = sheet.getRow(rowIndex);
                    if (destinationRow == null) {
                        destinationRow = sheet.createRow(rowIndex);
                    }
                    Cell destinationCell = destinationRow.createCell(destinationColumnIndex);
                    destinationCell.setCellValue( "The " + cellValue +" shown on Live monitoring tab will be correct");
                }
            }
        }
    }

}
