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
     * and value is the new cell value. This method also handles file copying from template.
     * 
     * @param templatePath Path to the template Excel file
     * @param outputPath Path where the updated Excel file should be saved
     * @param cellUpdates Map of cell addresses to values
     * @return Number of cells updated
     * @throws IOException If file operations fail
     */
    public static int updateExcelCells(String templatePath, String outputPath, Map<String, Object> cellUpdates) throws IOException {
        System.out.println("Starting Excel update from template: " + templatePath + " to output: " + outputPath);
        
        // Resolve template path - check if it's a classpath resource or absolute path
        java.io.InputStream templateInputStream = null;
        java.io.File templateFile = new java.io.File(templatePath);
        
        if (templateFile.exists()) {
            // Template exists as file path
            templateInputStream = new java.io.FileInputStream(templateFile);
            System.out.println("Using template file from filesystem: " + templateFile.getAbsolutePath());
        } else {
            // Try to load from classpath
            String resourcePath = templatePath.startsWith("src/main/resources/") ? 
                templatePath.substring("src/main/resources/".length()) : templatePath;
            templateInputStream = ExcelUpdater.class.getClassLoader().getResourceAsStream(resourcePath);
            if (templateInputStream == null) {
                throw new java.io.FileNotFoundException("Template file not found: " + templatePath + " (also tried as classpath resource: " + resourcePath + ")");
            }
            System.out.println("Using template file from classpath: " + resourcePath);
        }
        
        // Resolve output path - ensure it's in the project structure
        java.io.File outputFile = new java.io.File(outputPath);
        if (!outputFile.isAbsolute()) {
            // For relative paths, resolve them relative to the project root
            String projectRoot = System.getProperty("user.dir");
            // If we're running from the runtime directory, find the actual project
            if (projectRoot.contains("mule-enterprise-standalone")) {
                // We're in the Mule runtime, need to find the actual project directory
                String actualProjectPath = "/Users/sarvarth.bhatnagar/acb/rbc_excel/excel-update-integration";
                outputFile = new java.io.File(actualProjectPath, outputPath);
            } else {
                // We're in the project directory
                outputFile = new java.io.File(projectRoot, outputPath);
            }
        }
        String resolvedOutputPath = outputFile.getAbsolutePath();
        System.out.println("Resolved output path: " + resolvedOutputPath);
        
        // Create output directory if it doesn't exist
        java.io.File outputDir = outputFile.getParentFile();
        if (outputDir != null && !outputDir.exists()) {
            outputDir.mkdirs();
            System.out.println("Created output directory: " + outputDir.getAbsolutePath());
        }
        
        // Copy template to output location using input stream
        try (java.io.FileOutputStream fos = new java.io.FileOutputStream(resolvedOutputPath)) {
            byte[] buffer = new byte[8192];
            int bytesRead;
            while ((bytesRead = templateInputStream.read(buffer)) != -1) {
                fos.write(buffer, 0, bytesRead);
            }
        } finally {
            if (templateInputStream != null) {
                templateInputStream.close();
            }
        }
        System.out.println("Template file copied to: " + resolvedOutputPath);
        System.out.println("Starting Excel update for file: " + resolvedOutputPath);
        
        int updatedCount = 0;
        
        try (FileInputStream fis = new FileInputStream(resolvedOutputPath);
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
            try (FileOutputStream fos = new FileOutputStream(resolvedOutputPath)) {
                workbook.write(fos);
            }
            
            System.out.println("Successfully updated " + updatedCount + " cells in file: " + resolvedOutputPath);
            
        } catch (IOException e) {
            System.err.println("IO error while updating Excel file " + resolvedOutputPath + ": " + e.getMessage());
            throw e;
        } catch (Exception e) {
            System.err.println("Unexpected error while updating Excel file " + resolvedOutputPath + ": " + e.getMessage());
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
