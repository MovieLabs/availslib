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

    public AvailSS(String file, Logger logger, boolean exitOnError, boolean cleanupData) {
        this.file = file;
        this.logger = logger;
        this.exitOnError = exitOnError;
        this.cleanupData = cleanupData;
        sheets = new ArrayList<AvailsSheet>();
    }

    /**
     * Add a sheet from an Excel spreadsheet
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
        
        AvailsSheet as = new AvailsSheet(this, sheetName);
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
        wb.close();
        sheets.add(as);
        return as;
    }

    public Logger getLogger() {
        return logger;
    }

    public boolean getExitOnError() {
        return exitOnError;
    }

    public boolean getCleanupData() {
        return cleanupData;
    }

    /**
     * Dump raw contents of specified sheet
     * @param sheetName name of the sheet to dump
     */
    public void dump(String sheetName) {
        for (AvailsSheet s : sheets) {
            if (s.getName().equals(sheetName)) {
                int i = 0;
                for (SheetRow sr : s.getRows()) {
                    System.out.print("row " + i++ + "=[");
                    for (String cell : sr.getFields()) {
                        System.out.print("|" + cell );
                    }
                    System.out.println("]");
                }
            }
        }
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
