package org.example;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;

public class MonitorTestCases {
    public  static int startRow = 2; // Row index to start reading (zero-based)
    public  static int endRow = 56; // Row index to end reading (zero-based)
    public static void main(String[] args) throws IOException {
        String deviceName = "Mersive Solstice Pod";

        String projectPath = System.getProperty("user.dir");
        String importFilePath = projectPath + File.separator + "importFile" + File.separator + "Mersive Solstice Pod.xlsx";
        FileInputStream file = new FileInputStream(importFilePath);
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);

        try {
            writeTcTitle(sheet);
            writeTestStep(deviceName, sheet);
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
        int destinationColumnIndex = 5; // Index of the destination column (zero-based)

        // Read data

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
    private static void writeTestStep(String deviceName, Sheet sheet) throws IOException {


        int sourceColumnIndex = 2; // Index of the source column (zero-based)
        int destinationColumnIndex = 7; // Index of the destination column (zero-based)

        // Read data
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
                    destinationCell.setCellValue( "1. Login Symphony \n2.Go to " + deviceName + " device under test \n3. Go to Extended properties tab of the device \n4. Check the " + cellValue + " value \n5. Open web UI of " + deviceName + " on another browser \n6. Go to Licensing  tab \n7.Verify the result");
                }
            }
        }
    }
    private static void writeTestExpectedResult(Sheet sheet) throws IOException {

        int sourceColumnIndex = 2; // Index of the source column (zero-based)
        int typeColumnIndex = 3;
        int destinationColumnIndex = 9; // Index of the destination column (zero-based)

        // Read data
        for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
            Row sourceRow = sheet.getRow(rowIndex);
            if (sourceRow != null) {
                Cell sourceCell = sourceRow.getCell(sourceColumnIndex);
                Cell typeCell = sourceRow.getCell(typeColumnIndex);
                if (sourceCell != null) {
                    // Retrieve the value from the source cell
                    String cellValue = getStringCellValue(sourceCell);
                    String typeValue = getStringCellValue(typeCell);
                    System.out.println(typeValue);

                    // Write the value to the destination column
                    Row destinationRow = sheet.getRow(rowIndex);
                    Cell destinationCell = destinationRow.createCell(destinationColumnIndex);
                    if (destinationRow == null) {
                        destinationRow = sheet.createRow(rowIndex);
                    }
                    if (typeValue.startsWith("button") || typeValue.startsWith("Button")) {
                        destinationCell.setCellValue("The " + cellValue + " should be a button");
                    }else if (typeValue.startsWith("switch button") || typeValue.startsWith("Switch button")) {
                        destinationCell.setCellValue("The " + cellValue + " is a switch button shown on Live Monitoring matches the " + cellValue + " shown on the device web page.");
                    }else if (typeValue.startsWith("dropdown") || typeValue.startsWith("Dropdown")){
                        destinationCell.setCellValue("The " + cellValue + " is a dropdown list that contains " + typeValue.substring(10));
                    }else if (typeValue.startsWith("slider") || typeValue.startsWith("Slider")) {
                        destinationCell.setCellValue("The " + cellValue + " should be a slider with range is " + typeValue.substring(8));
                    }else if (typeValue.isEmpty()){
                        destinationCell.setCellValue( "The " + cellValue +" shown on Live monitoring tab will be correct");
                    }
                }
            }
        }
    }

}
