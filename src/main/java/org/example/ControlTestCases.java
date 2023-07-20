package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ControlTestCases {
    public static int startRow = 1; // Row index to start reading (zero-based)
    public static int endRow = 70; // Row index to end reading (zero-based)
    public static String deviceName = "Mersive Solstice Pod";

    public static void main(String[] args) throws IOException {

        String projectPath = System.getProperty("user.dir");
        String importFilePath = projectPath + File.separator + "importFile" + File.separator + "Mersive Solstice Pod.xlsx";
        FileInputStream file = new FileInputStream(importFilePath);
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheetMonitor = workbook.getSheetAt(0); // Sheet index(zero-based)
        Sheet sheetControl = workbook.getSheetAt(1); // Sheet index(zero-based)


        try {
            writeTcTitle(sheetControl, sheetMonitor);
            writeTestStep(sheetControl);
            writeTestExpectedResult(sheetControl, sheetMonitor);

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


    private static void writeTcTitle(Sheet controlSheet, Sheet monitorSheet) throws IOException {

        int propertyColumnIndex = 2; // Index of the source column (zero-based)
        int optionColumnIndex = 3; // Index of the source column (zero-based)
        int destinationColumnIndex = 4; // Index of the destination column (zero-based)
        String propertyType = null;

        // Read data
        for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
            Row sourceRow = controlSheet.getRow(rowIndex);
            if (sourceRow != null) {
                Cell propertyCell = sourceRow.getCell(propertyColumnIndex);
                Cell optionCell = sourceRow.getCell(optionColumnIndex);
                if (propertyCell != null) {
                    // Retrieve the value from the source cell
                    String propertyValue = getStringCellValue(propertyCell);
                    String optionValue = getStringCellValue(optionCell);
                    propertyType = getTypeFromMonitorSheet(monitorSheet, propertyValue);

                    // Write the value to the destination column
                    Row destinationRow = controlSheet.getRow(rowIndex);
                    Cell destinationCell = destinationRow.createCell(destinationColumnIndex);
//                    System.out.println(propertyValue);
//                    System.out.println(optionValue);
//                    System.out.println(propertyType);

                    if ((optionValue.contains("device") || optionValue.contains("Device")) && (propertyType.startsWith("button") || propertyType.startsWith("Button"))) {
                        destinationCell.setCellValue("Validate the user can click on " + propertyValue + "  button for " + deviceName + "  from real device");
                    }
                    else if ((optionValue.contains("symphony") || optionValue.contains("Symphony")) && (propertyType.startsWith("button") || propertyType.startsWith("Button"))) {
                        destinationCell.setCellValue("Validate the user can click on " + propertyValue + "  button for " + deviceName + " from device management");
                    }
                    else if ((optionValue.contains("device") || optionValue.contains("Device")) && (propertyType.startsWith("switch button") || propertyType.startsWith("Switch button"))) {
                        destinationCell.setCellValue("Validate the user can turn ON/OFF " + propertyValue + " of " + deviceName + " from real device");
                    }
                    else if ((optionValue.contains("symphony") || optionValue.contains("Symphony")) && (propertyType.contains("switch button") || propertyType.contains("Switch button"))) {
                        destinationCell.setCellValue("Validate the user can turn ON/OFF " + propertyValue + " of " + deviceName + " from device management");
                    }
                    else if ((optionValue.contains("device") || optionValue.contains("Device")) && (propertyType.isEmpty() || propertyType.contains("dropdown") || propertyType.contains("Dropdown"))) {
                        destinationCell.setCellValue("Validate the user can change " + propertyValue + " of " + deviceName + " from real device");
                    }
                    else if ((optionValue.contains("symphony") || optionValue.contains("Symphony")) && (propertyType.isEmpty() || propertyType.contains("dropdown") || propertyType.contains("Dropdown"))) {
                        destinationCell.setCellValue("Validate the user can change " + propertyValue + " of " + deviceName + " from device management");
                    }

                }
            }
        }
    }

    private static void writeTestStep(Sheet sheet) throws IOException {

        int sourceColumnIndex = 2; // Index of the source column (zero-based)
        int titleColumnIndex = 4; // Index of the source column (zero-based)
        int testStepColumnIndex = 6; // Index of the destination column (zero-based)

        // Read data
        for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
            Row sourceRow = sheet.getRow(rowIndex);
            if (sourceRow != null) {
                Cell sourceCell = sourceRow.getCell(sourceColumnIndex);
                Cell titleCell = sourceRow.getCell(titleColumnIndex);
                if (sourceCell != null) {
                    // Retrieve the value from the source cell
                    String cellValue = getStringCellValue(sourceCell);
                    String titleValue = getStringCellValue(titleCell);

                    // Write the value to the destination column
                    Row testStepRow = sheet.getRow(rowIndex);
                    Cell testStepCell = testStepRow.createCell(testStepColumnIndex);
                    if (titleValue.contains("ON/OFF") && titleValue.contains("real device")) {
                        testStepCell.setCellValue("On " + deviceName + " web UI:\n" +
                                "1. Go to Config page\n" +
                                "2. Go to Network\n" +
                                "3. Switch the " + cellValue + " button to OFF \n" +
                                "On Symphony:\n" +
                                "4. Check the " + cellValue + " value of the device \n" +
                                "On " + deviceName + " web UI: \n" +
                                "5. Switch the " + cellValue + "button to ON\n" +
                                " On Symphony:\n" +
                                "6. Check the " + cellValue + "of the device");
                    } else if (titleValue.contains("ON/OFF") && titleValue.contains("device management")) {
                        testStepCell.setCellValue("On Symphony:\n" +
                                "1. Go to " + deviceName + " device under test\n" +
                                "2. Go to monitor tab of the device\n" +
                                "3. Switch the " + cellValue + " button to OFF \n" +
                                "On " + deviceName + " web UI:\n" +
                                "4. Check the " + cellValue + " value of the device \n" +
                                "On Symphony:\n" +
                                "5. Switch the " + cellValue + "button to ON\n" +
                                "6. ON the " + cellValue + "\n\n" +
                                "7.Check the " + cellValue + "of the device");
                    } else if (titleValue.contains("can change") && titleValue.contains("real device")) {
                        testStepCell.setCellValue("On " + deviceName + " web UI:\n" +
                                "1. Go to Config page\n" +
                                "2. Go to Network\n" +
                                "3. Change value of " + cellValue + "\n" +
                                "On Symphony:\n" +
                                "4. Check the " + cellValue + " value of the device ");
                    } else if (titleValue.contains("can change") && titleValue.contains("device management")) {
                        testStepCell.setCellValue("On Symphony:\n" +
                                "1. Go to " + deviceName + " device under test\n" +
                                "2. Go to monitor tab of the device\n" +
                                "3. Change value of " + cellValue + "\n" +
                                "On " + deviceName + " web UI:\n" +
                                "4. Check the " + cellValue + " value of the device");
                    } else if (titleValue.contains("click on") && titleValue.contains("real device")) {
                        testStepCell.setCellValue("On " + deviceName + " web UI:\n" +
                                "1. Go to Config page\n" +
                                "2. Go to Network\n" +
                                "3. Click on " + cellValue + " button\n" +
                                "4. Validate the result");
                    } else if (titleValue.contains("click on") && titleValue.contains("device management")) {
                        testStepCell.setCellValue("On Symphony:\n" +
                                "1. Go to " + deviceName + " device under test\n" +
                                "2. Go to monitor tab of the device\n" +
                                "3. Change value of " + cellValue + "\n" +
                                "4. Validate the result");
                    }

                }
            }
        }
    }

    private static void writeTestExpectedResult(Sheet controlSheet, Sheet monitorSheet) throws IOException {

        int sourceColumnIndex = 2; // Index of the source column (zero-based)
        int typeColumnIndex = 3; // Index of the source column (zero-based)
        int titleColumnIndex = 4; // Index of the source column (zero-based)
        int expectedResultColumnIndex = 8; // Index of the destination column (zero-based)

        // Read data
        for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
            Row sourceRow = controlSheet.getRow(rowIndex);

            if (sourceRow != null) {
                Cell sourceCell = sourceRow.getCell(sourceColumnIndex);
                Cell titleCell = sourceRow.getCell(titleColumnIndex);

                if (sourceCell != null) {
                    // Retrieve the value from the source cell
                    String cellValue = getStringCellValue(sourceCell);
                    String titleValue = getStringCellValue(titleCell);
                    String propertyType = getTypeFromMonitorSheet(monitorSheet, cellValue);

                    // Write the value to the destination column
                    Row expectedResultRow = controlSheet.getRow(rowIndex);
                    Cell expectedResultCell = expectedResultRow.createCell(expectedResultColumnIndex);

                    if (titleValue.contains("ON/OFF") && titleValue.contains("real device")) {
                        expectedResultCell.setCellValue("The " + cellValue + " is a switch button and must be updated on the Symphony correctly");
                    } else if (titleValue.contains("ON/OFF") && titleValue.contains("device management")) {
                        expectedResultCell.setCellValue("The " + cellValue + " is a switch button and must be updated on the device correctly");
                    }
                    else if (titleValue.contains("can change") && titleValue.contains("real device")) {
                        expectedResultCell.setCellValue("The " + cellValue + " must be updated on the Symphony correctly");
                    } else if (titleValue.contains("can change") && titleValue.contains("device management")) {
                        expectedResultCell.setCellValue("The " + cellValue + " must be updated on the device correctly");
                    }
                    else if (titleValue.contains("click on") && titleValue.contains("real device")) {
                        expectedResultCell.setCellValue("The " + cellValue + " button works well");
                    } else if (titleValue.contains("click on") && titleValue.contains("device management")) {
                        expectedResultCell.setCellValue("The " + cellValue + " button works well");
                    }
                }
            }
        }
    }

    private static String getTypeFromMonitorSheet(Sheet monitorSheet, String property) {
        int startRow = 1;
        int endRow = 56;
        int propertyColumnIndex = 2;
        int typeColumnIndex = 3;
        String typeValue = "";
        for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
            Row propertyRow = monitorSheet.getRow(rowIndex);
            if (propertyRow != null) { // Add a check for null propertyRow
                Cell cell = propertyRow.getCell(propertyColumnIndex);
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    String propertyValue = cell.getStringCellValue();

                    // Check if the value of column X is "A"
                    if (propertyValue.equals(property)) {
                        // Get the value of column X + 1 for the current row
                        Row typeRow = monitorSheet.getRow(rowIndex);
                        if (typeRow != null) { // Add a check for null typeRow
                            Cell typeCell = typeRow.getCell(typeColumnIndex);

                            // Perform further processing with the value in column X + 1 (if needed)
                            if (typeCell != null && typeCell.getCellType() == CellType.STRING) {
                                String type = typeCell.getStringCellValue();
                                typeValue = type;
                            } else {
                                typeValue = "";
                            }
                        } else {
                            typeValue = "";
                        }
                    }
                }
            }
        }
        return typeValue;
    }
}

