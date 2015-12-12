package com.movielabs.avails;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.*;

import org.apache.logging.log4j.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Represents an Excel spreadsheet comprising multiple individual
 * sheets, each of which are represented by an AvailsSheet object
 */
public class AvailSS {
    private String file;
    private ArrayList<AvailsSheet> sheets;
    private static Logger logger;
    private boolean exitOnError;
    private boolean cleanupData;

    /**
     * Create a Spreadsheet object
     * @param file name of the Excel Spreadsheet file
     * @param logger a log4j logger object
     * @param exitOnError true if validation errors should cause immediate failure
     * @param cleanupData true if minor validation errors should be auto-corrected
     */
    public AvailSS(String file, Logger logger, boolean exitOnError, boolean cleanupData) {
        this.file = file;
        this.logger = logger;
        this.exitOnError = exitOnError;
        this.cleanupData = cleanupData;
        sheets = new ArrayList<AvailsSheet>();
    }

    /**
     * Add a sheet from an Excel spreadsheet to a spreatsheet object
     * @param wb an Apache POI workbook object
     * @param sheet an Apache POI sheet object
     * @return created sheet object
     */
    private AvailsSheet addSheetHelper(Workbook wb, Sheet sheet) throws Exception {
        AvailsSheet as = new AvailsSheet(this, sheet.getSheetName());

        for (Row row : sheet) {
            int len = 0;
            for (Cell cell : row) len++;
            if (len <= 0)
                break;
            String[] fields = new String[len];
            for (Cell cell : row) {
                int idx = cell.getColumnIndex();
                fields[idx] = cell.toString();
            }
            if (as.isAvail(fields))
                as.addRow(fields);
        }
        sheets.add(as);
        return as;
    }

    /**
     * Add a sheet from an Excel spreadsheet to a spreadsheet object
     * @param sheetName name of the sheet to add
     * @return created sheet object
     * @throws IllegalArgumentException if the sheet does not exist in the Excel spreadsheet
     * @throws Exception other error conditions may also throw exceptions
     */
    public AvailsSheet addSheet(String sheetName) throws Exception {
        Workbook wb = new XSSFWorkbook(new FileInputStream(file));
        Sheet sheet = wb.getSheet(sheetName);
        if (sheet == null) {
            wb.close();
            throw new IllegalArgumentException(file + ":" + sheetName + " not found");
        }
        AvailsSheet as = addSheetHelper(wb, sheet);
        wb.close();
        return as;
    }

    /**
     * Add a sheet from an Excel spreadsheet to a spreadsheet object
     * @param sheetNumber zero-based index of sheet to add
     * @return created sheet object
     * @throws IllegalArgumentException if the sheet does not exist in the Excel spreadsheet
     * @throws Exception other error conditions may also throw exceptions
     */
    public AvailsSheet addSheet(int sheetNumber) throws Exception {
        Workbook wb = new XSSFWorkbook(new FileInputStream(file));

        Sheet sheet;
        try {
            sheet = wb.getSheetAt(sheetNumber);
        } catch (IllegalArgumentException e) {
            wb.close();
            throw new IllegalArgumentException(file + ": sheet number " + sheetNumber + " not found");
        }
        AvailsSheet as = addSheetHelper(wb, sheet);
        wb.close();
        return as;
    }

    /**
     * Get the logging object
     * @return Logger for this instance
     */
    public Logger getLogger() {
        return logger;
    }

    /**
     * Get the error handling option
     * @return true if exiting on encountering an invalid cell
     */
    public boolean getExitOnError() {
        return exitOnError;
    }

    /**
     * Get the data cleaning option
     * @return true minor validation errors will be fixed up
     */
    public boolean getCleanupData() {
        return cleanupData;
    }

    /**
     * Dump raw contents of specified sheet
     * @param sheetName name of the sheet to dump
     * @throws Exception if any error is encountered (e.g. non-existant or corrupt file)
     */
    public void dumpSheet(String sheetName) throws Exception {
        boolean foundSheet = false;
        for (AvailsSheet s : sheets) {
            if (s.getName().equals(sheetName)) {
                int i = 0;
                foundSheet = true;
                for (SheetRow sr : s.getRows()) {
                    System.out.print("row " + i++ + "=[");
                    for (String cell : sr.getFields()) {
                        System.out.print("|" + cell );
                    }
                    System.out.println("]");
                }
            }
        }
        if (!foundSheet)
            throw new IllegalArgumentException(file + ":" + sheetName + " not found");
    }

    /**
     * Dump the contents (sheet-by-sheet) of an Excel spreadsheet
     * @param file name of the Excel .xlsx spreadsheet
     * @throws Exception if any error is encountered (e.g. non-existant or corrupt file)
     */
    public static void dumpFile(String file) throws Exception {
        Workbook wb = new XSSFWorkbook(new FileInputStream(file));
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet sheet = wb.getSheetAt(i);
            System.out.println("Sheet <" + wb.getSheetName(i) + ">");
            for (Row row : sheet) {
                System.out.println("rownum: " + row.getRowNum());
                for (Cell cell : row) {
                    System.out.println("   | " + cell.toString());
                }
            }
        }
        wb.close();
    }
}
