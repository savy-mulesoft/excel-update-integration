package com.mulesoft.excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Excel file updater using Apache POI library
 * Supports updating individual cells in XLSX/XLSM files
 */
public class ExcelUpdater {
    
    /**
     * Updates Excel file with key-value pairs where key is cell address (e.g., "A1", "B2")
     * and value is the new cell value
     * 
     * @param filePath Path to the Excel file
     * @param cellUpdates Map of cell addresses to values
     * @return Number of cells updated
     * @throws IOException If file operations fail
     */
    public static int updateExcelCells(String filePath, Map<String, Object> cellUpdates) throws IOException {
        System.out.println("Starting Excel update for file: " + filePath);
        
        int updatedCount = 0;
        
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            
            // Get the first sheet (Sheet1)
            Sheet sheet = workbook.getSheetAt(0);
            if (sheet == null) {
                throw new IllegalArgumentException("No sheets found in the Excel file");
            }
            
            System.out.println("Working with sheet: " + sheet.getSheetName());
            
            // Update each cell
            for (Map.Entry<String, Object> entry : cellUpdates.entrySet()) {
                String cellAddress = entry.getKey();
                Object cellValue = entry.getValue();
                
                try {
                    // Parse cell address (e.g., "A1" -> row=0, col=0)
                    CellReference cellRef = new CellReference(cellAddress);
                    int rowIndex = cellRef.getRow();
                    int colIndex = cellRef.getCol();
                    
                    // Get or create row
                    Row row = sheet.getRow(rowIndex);
                    if (row == null) {
                        row = sheet.createRow(rowIndex);
                    }
                    
                    // Get or create cell
                    Cell cell = row.getCell(colIndex);
                    if (cell == null) {
                        cell = row.createCell(colIndex);
                    }
                    
                    // Set cell value based on type
                    setCellValue(cell, cellValue);
                    
                    System.out.println("Updated cell " + cellAddress + " with value: " + cellValue);
                    updatedCount++;
                    
                } catch (Exception e) {
                    System.err.println("Failed to update cell " + cellAddress + ": " + e.getMessage());
                    throw new RuntimeException("Failed to update cell " + cellAddress, e);
                }
            }
            
            // Save the workbook back to file
            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }
            
            System.out.println("Successfully updated " + updatedCount + " cells in file: " + filePath);
            
        } catch (IOException e) {
            System.err.println("IO error while updating Excel file " + filePath + ": " + e.getMessage());
            throw e;
        } catch (Exception e) {
            System.err.println("Unexpected error while updating Excel file " + filePath + ": " + e.getMessage());
            throw new RuntimeException("Failed to update Excel file", e);
        }
        
        return updatedCount;
    }
    
    /**
     * Sets the cell value based on the object type
     */
    private static void setCellValue(Cell cell, Object value) {
        if (value == null) {
            cell.setBlank();
        } else if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else {
            // Convert to string for other types
            cell.setCellValue(value.toString());
        }
    }
    
    /**
     * Validates if a cell address is in correct format (e.g., A1, B2, etc.)
     */
    public static boolean isValidCellAddress(String cellAddress) {
        try {
            new CellReference(cellAddress);
            return true;
        } catch (Exception e) {
            return false;
        }
    }
}
