package com.movielabs.avails;

import java.io.FileInputStream;
import java.util.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SS {

    public enum COL {
        DisplayName ("DisplayName"),
        StoreLanguage ("StoreLanguage"),
        Territory ("Territory"),
        WorkType ("WorkType"),
        EntryType ("EntryType"),
        TitleInternalAlias ("TitleInternalAlias"),
        TitleDisplayUnlimited ("TitleDisplayUnlimited"),
        LocalizationType ("LocalizationOffering"),
        LicenseType ("LicenseType"),
        LicenseRightsDescription ("LicenseRightsDescription"),
        FormatProfile ("FormatProfile"),
        Start ("StartCondition"),
        End ("EndCondition"),
        PriceType ("PriceType"),
        PriceValue ("PriceValue"),
        SRP ("SRP"),
        Description ("Description"),
        OtherTerms ("OtherTerms"),
        OtherInstructions ("OtherInstructions"),
        ContentID ("ContentID"),
        ProductID ("EditEIDR-S"),
        EncodeID ("EncodeID"),
        AvailID ("AvailID"),
        Metadata ("Metadata"),
        AltID ("Identifier"),
        SuppressionLiftDate ("AnnounceDate"),
        SpecialPreOrderFulfillDate ("SpecialPreOrderFulfillDate"),
        ReleaseYear ("ReleaseDate"),
        ReleaseHistoryOriginal ("Date"),
        ReleaseHistoryPhysicalHV ("Date"),
        ExceptionFlag ("ExceptionFlag"),
        RatingSystem ("RatingSystem"),
        RatingValue ("RatingValue"),
        RatingReason ("RatingReason"),
        RentalDuration ("RentalDuration"),
        WatchDuration ("WatchDuration"),
        CaptionIncluded ("USACaptionsExemptionReason"),
        CaptionExemption ("USACaptionsExemptionReason"),
        Any ("Any"),
        ContractID ("ContractID"),
        ServiceProvider ("ServiceProvider"),
        TotalRunTime ("RunLength"),
        HoldbackLanguage ("HoldbackLanguage"),
        HoldbackExclusionLanguage ("HoldbackExclusionLanguage"); // 43

        private final String name;

        private COL(String s) {
            name = s;
        }

        public boolean equalsName(String otherName) {
            return (otherName == null) ? false : name.equals(otherName);
        }

        public String toString() {
            return this.name;
        }
    }


    private String sheet;
    private ArrayList<Object> rows;

    private boolean isAvail(String[] row) {
        String t = row[COL.Territory.ordinal()];
        boolean ret = (t != null) && 
            !(t.equals("AvailTrans") || t.equals("Territory") || t.substring(0, 2).equals("//"));
        return ret;
    }

    public void toXML(boolean strict) throws Exception
    {
        XMLGen xml = new XMLGen(strict);
        xml.makeXML(rows);
    }

    public SS(String file, String sheetName) throws Exception {
        sheet = sheetName;
        Workbook wb = new XSSFWorkbook(new FileInputStream(file));
        Sheet sheet = wb.getSheet(sheetName);
        if (sheet == null) {
            wb.close();
            throw new IllegalArgumentException(file + ":" + sheetName + " not found");
        }
        System.out.println("Sheet <" + sheetName + ">");
        int nrows = 0;
        for (Row row : sheet) nrows++;
        rows = new ArrayList<Object>(nrows);
        
        for (Row row : sheet) {
            int last = COL.HoldbackExclusionLanguage.ordinal();
            String[] rvals = new String[last + 1];
            for (Cell cell : row) {
                int idx = cell.getColumnIndex();
                if (idx >= 0 && idx < COL.values().length) {
                    rvals[idx] = cell.toString();
                } else {
                    break;
                }
            }
            if (isAvail(rvals))
                rows.add(rvals);
        }
        wb.close();
    }

    public void dump() {
        int i = 0;
        for (Object r : rows) {
            String[] row = (String[]) r;
            System.out.print("row " + i++ + "=[");
            for (String cell : row) {
                System.out.print("|" + cell );
            }
            System.out.println("]");
        }
    }

    public static void dump(String file) throws Exception {
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
