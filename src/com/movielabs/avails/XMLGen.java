package com.movielabs.avails;

import java.io.*;
import java.util.*;
import java.text.ParseException;
import java.lang.NumberFormatException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.*;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.w3c.dom.*;

import com.movielabs.avails.SS.COL;
import org.apache.logging.log4j.*;

public class XMLGen {
    ArrayList<String> LRD = new ArrayList<String>(Arrays.asList("New Release",
                                                                "Library",
                                                                "Mega-Library",
                                                                "Priority Library",
                                                                "DD-Theatrical",
                                                                "Pre-Theatrical",
                                                                "DD-DVD",
                                                                "Early EST",
                                                                "Preorder EST",
                                                                "Early VOD",
                                                                "Preorder VOD",
                                                                "DTV",
                                                                "Next Day TV",
                                                                "Season Only",
                                                                "Free"));

    private static final String MISSING = "---missing---";
    private Document dom;
    private Element root;
    private boolean strict = false;

    public XMLGen(boolean strict) {
        this.strict = strict;
    }

    private Node mALID(String availID) {
        if (availID.equals(""))
            availID = MISSING;
        Text tmp = dom.createTextNode(availID);
        Element ALID = dom.createElement("ALID");
        ALID.appendChild(tmp);
        return ALID;
    }

    /**
     * Creates an Avails/Disposition XML node
     * @param entryType string (controlled vocabulary) indicating whether this Avail is new, and update, or a deletion
     * @return the created XML node
     */
    private Element mDisposition(String entryType) throws Exception {
        Comment comment = null;

        if (!(entryType.equals("Full Extract") || entryType.equals("Full Delete"))) {
            if (strict) {
                throw new ParseException("invalid Disposition", 0);
            } else { // try to clean up
                Pattern pat = Pattern.compile("^\\s*full\\s+(extract|delete)\\s*$", Pattern.CASE_INSENSITIVE);
                Matcher m = pat.matcher(entryType);
                if (m.matches()) {
                    comment = dom.createComment("corrected from '" + entryType + "'");
                    if (m.group(1).equalsIgnoreCase("extract"))
                        entryType = "Full Extract";
                    else if (m.group(1).equalsIgnoreCase("delete"))
                        entryType = "Full Delete";
                    else
                        throw new ParseException("invalid Disposition", 0);
                } else {
                    throw new ParseException("invalid Disposition", 0);
                }
            }
        }
        Element disp = dom.createElement("Disposition");
        Element entry = dom.createElement("EntryType");
        Text tmp = dom.createTextNode(entryType);
        entry.appendChild(tmp);
        disp.appendChild(entry);
        if (comment != null)
            disp.appendChild(comment);
        return disp;
    }

    private Node mLicensor(String displayName) throws Exception {
        if (displayName.equals(""))
            throw new ParseException("missing DisplayName", 0);
        Element licensor = dom.createElement("Licensor");
        Element dname = dom.createElement("DisplayName");
        Text tmp = dom.createTextNode(displayName);
        dname.appendChild(tmp);
        licensor.appendChild(dname);

        return licensor;
    }
 
    private int yesorno(String s) {
        if (s.equals(""))
            return 0;
        Pattern pat = Pattern.compile("^\\s*y(?:es)?\\s*$", Pattern.CASE_INSENSITIVE); // Y, Yes, yes, yEs, etc.
        Matcher m = pat.matcher(s);
        if (m.matches()) {
            return 1;
        } else {
            pat = Pattern.compile("^\\s*n(?:o)?\\s*$", Pattern.CASE_INSENSITIVE);  // N, No, no, nO, etc.
            m = pat.matcher(s);
            if (m.matches())
                return 0;
            else
                return -1;
        }
    }

    private Node mException(String exceptionFlag) throws Exception {
        switch(yesorno(exceptionFlag)) {
        case 0:
            return null;
        case 1:
            Element eFlag = dom.createElement("ExceptionFlag");
            Text tmp = dom.createTextNode("true");
            eFlag.appendChild(tmp);
            return eFlag;
        default:
            throw new ParseException("invalid ExceptionFlag", 0);
        }
    }
 
    private Element mGenericElement(String[] row, COL element, boolean mandatory) throws Exception {
        String elementName = element.toString();
        String val = row[element.ordinal()];
        if (val.equals("")) {
            if (mandatory)
                throw new ParseException("missing required value on element: " + elementName, 0);
            else
                return null;
        }
        Element tia = dom.createElement(elementName);
        Text tmp = dom.createTextNode(val);
        tia.appendChild(tmp);
        return tia;
    }

    private Element mLocalizationType(String[] row, COL element, String[] cmt) throws Exception {
        String loc = row[element.ordinal()];

        cmt[0] = "";
        if (loc.equals(""))
            return null;
        if (!(loc.equals("sub") || loc.equals("dub") || loc.equals("subdub") || loc.equals("any"))) {
            if (strict) {
                throw new ParseException("invalid LocalizationOffering value: " + loc, 0);
            } else {
                Pattern pat = Pattern.compile("^\\s*(sub|dub|subdub|any)\\s*$", Pattern.CASE_INSENSITIVE);
                Matcher m = pat.matcher(loc);
                if (m.matches()) {
                    cmt[0] = "corrected from '" + loc + "'";
                    row[element.ordinal()] = m.group(1).toLowerCase();
                }
            }
        }
        return mGenericElement(row, COL.LocalizationType, false);
    }

    private Element mCaptionsExemptionReason(String[] row, String territory) throws Exception {
        //        if (territory.equals("US")) {  // check captions
            switch(yesorno(row[COL.CaptionIncluded.ordinal()])) {
            case 1:
                if (!row[COL.CaptionExemption.ordinal()].equals(""))
                    throw new ParseException("CaptionExemption specified without CaptionIncluded", 0);
                else
                    return null;
            case 0:
                if (row[COL.CaptionExemption.ordinal()].equals(""))
                    throw new ParseException("CaptionExemption not specified", 0);
                int exemption = normalizeInt(row[COL.CaptionExemption.ordinal()]);
                if (exemption < 1 || exemption > 6)
                    throw new ParseException("Invalid CaptionExamption Value: " + exemption, 0);
                Element capex = dom.createElement(COL.CaptionExemption.toString());
                Text tmp = dom.createTextNode(Integer.toString(exemption));
                capex.appendChild(tmp);
                return capex;
            default:
                throw new ParseException("CaptionExemption specified without CaptionIncluded", 0);
            }
            // } else {
            // if (row[COL.CaptionIncluded.ordinal()].equals("") && row[COL.CaptionExemption.ordinal()].equals(""))
            //     return null;
            // else
            //     throw new ParseException("CaptionIncluded/Exemption should only be specified in US", 0);
            // }
    }

    private Element mRunLength(String[] row, COL element) throws Exception {
        String elementName = element.toString();
        String val = row[element.ordinal()];
        if (val.equals(""))
            return null;
        Pattern dur = Pattern.compile("^\\s*(\\d{1,2}):(\\d{1,2}):(\\d{1,2})\\s*$");
        Matcher m = dur.matcher(val);
        if (m.matches()) {
            int hour = normalizeInt(m.group(1));
            int min = normalizeInt(m.group(2));
            int sec = normalizeInt(m.group(3));
            if (min > 59 || sec > 59)
                throw new ParseException("invalid duration string " + val, 0);
            String d =  String.format("P%dH%dM%dS", hour, min, sec);
            Element ret = dom.createElement(elementName);
            Text tmp = dom.createTextNode(d);
            ret.appendChild(tmp);
            return ret;
        } else {
            throw new ParseException("invalid duration string (verify Excel cell format is 'Text'): " + val, 0);
        }
    }

    private String normalizeEIDR(String s) {
        Pattern eidr = Pattern.compile("^\\s*(?:10\\.5240/)?((?:(?:\\p{XDigit}){4}-){5}\\p{XDigit})\\s*$");
        Matcher m = eidr.matcher(s);
        if (m.matches()) {
            return m.group(1).toUpperCase();
        } else {
            return null;
        }
    }

    private String normalizeYear(String s) {
        Pattern eidr = Pattern.compile("^\\s*(\\d{4})(?:\\.0)?\\s*$");
        Matcher m = eidr.matcher(s);
        if (m.matches()) {
            return m.group(1);
        } else {
            return null;
        }
    }

    private int normalizeInt(String s) throws Exception {
        Pattern eidr = Pattern.compile("^\\s*(\\d+)(?:\\.0)?\\s*$");
        Matcher m = eidr.matcher(s);
        if (m.matches()) {
            return Integer.parseInt(m.group(1));
        } else {
            throw new NumberFormatException(s);
        }
    }

    private String normalizeDate(String s) throws Exception {
        final int[] dim = {31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31};
        int year=-1, month=-1, day=-1;
        Pattern date = Pattern.compile("^\\s*(\\d{4})-(\\d{1,2})-(\\d{1,2})\\s*$");
        Matcher m = date.matcher(s);
        if (m.matches()) { // try yyyy-mm-dd
            year = Integer.parseInt(m.group(1));
            month = Integer.parseInt(m.group(2));
            day = Integer.parseInt(m.group(3));
        } else { // try dd-mmm-yyyy
            date = Pattern.compile("^\\s*(\\d{1,2})-(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)-(\\d{4})\\s*$",
                                   Pattern.CASE_INSENSITIVE);
            m = date.matcher(s);
            if (m.matches()) {
                year = Integer.parseInt(m.group(3));
                switch(m.group(2).toLowerCase()) {
                case "jan": month =  1; break;
                case "feb": month =  2; break;
                case "mar": month =  3; break;
                case "apr": month =  4; break;
                case "may": month =  5; break;
                case "jun": month =  6; break;
                case "jul": month =  7; break;
                case "aug": month =  8; break;
                case "sep": month =  9; break;
                case "oct": month = 10; break;
                case "nov": month = 11; break;
                case "dec": month = 12; break;
                }
                day = Integer.parseInt(m.group(1));
            } else {
                return null;
            }
        }
        boolean badDate = year < 1850;
        badDate |= month < 1 || month > 12;
        badDate |= day < 1;
        if (month == 2) {  // February
            if ((year % 4) == 0) { // leap year: fails in year 2400
                if ((year % 100) == 0) {
                    badDate |= day > 28; // first-order exception to leap year rule
                } else {
                    badDate |= day > 29; // normal leap year
                }
            } else {
                badDate |= day > dim[1]; // non-leap year
            }
        } else {
            badDate |= day > dim[month-1];
        }
        if (badDate)
            return null;
        else
            return String.format("%04d-%02d-%02d", year, month, day);
    }

    private void mMovieAsset(String[] row, Element asset) throws Exception {
        Element e;
        String contentID = row[COL.ContentID.ordinal()];
        if (contentID.equals(""))
            contentID = MISSING;
        Attr attr = dom.createAttribute(COL.ContentID.name());
        attr.setValue(contentID);
        asset.setAttributeNode(attr);
        Element metadata = dom.createElement("Metadata");
        String territory = row[COL.Territory.ordinal()];

        // TitleInternalAlias
        metadata.appendChild(mGenericElement(row, COL.TitleInternalAlias, true));
        // TitleDisplayUnlimited
        if ((e = mGenericElement(row, COL.TitleDisplayUnlimited, false)) != null)
            metadata.appendChild(e);
        // LocalizationType --> LocalizationOffering
        String[] cmt = new String[1];
        if ((e = mLocalizationType(row, COL.LocalizationType, cmt)) != null) {
            if (!cmt[0].equals("")) {
                Comment comment = dom.createComment(cmt[0]);
                metadata.appendChild(comment);
            }
            metadata.appendChild(e);
        }
        // ProductID --> EditEIDR-S
        if (!row[COL.ProductID.ordinal()].equals("")) { // optional field
            String productID = normalizeEIDR(row[COL.ProductID.ordinal()]);
            if (productID == null) {
                throw new ParseException("Invalid ProductID: " + row[COL.ProductID.ordinal()], 0);
            } else {
                row[COL.ProductID.ordinal()] = productID;
            }
            metadata.appendChild(mGenericElement(row, COL.ProductID, false));
        }
        // EncodeID --> EditEIDR-S
        if (!row[COL.EncodeID.ordinal()].equals("")) { // optional field
            String encodeID = normalizeEIDR(row[COL.EncodeID.ordinal()]);
            if (encodeID == null) {
                throw new ParseException("Invalid EncodeID: " + row[COL.EncodeID.ordinal()], 0);
            } else {
                row[COL.EncodeID.ordinal()] = encodeID;
            }
            metadata.appendChild(mGenericElement(row, COL.EncodeID, false));
        }
        // AltID --> AltIdentifier
        if (!row[COL.AltID.ordinal()].equals("")) { // optional
            Element altID = dom.createElement("AltIdentifier");
            Element cid = dom.createElement("Namespace");
            Text tmp = dom.createTextNode(MISSING);
            cid.appendChild(tmp);
            altID.appendChild(cid);
            altID.appendChild(mGenericElement(row, COL.AltID, false));
            Element loc = dom.createElement("Location");
            tmp = dom.createTextNode(MISSING);
            loc.appendChild(tmp);
            altID.appendChild(loc);
            metadata.appendChild(altID);
        }
        // ReleaseYear ---> ReleaseDate
        if (!row[COL.ReleaseYear.ordinal()].equals("")) { // optional
            String year = normalizeYear(row[COL.ReleaseYear.ordinal()]);
            if (year == null) {
                throw new ParseException("Invalid ReleaseYear: " + row[COL.ReleaseYear.ordinal()], 0);
            } else {
                row[COL.ReleaseYear.ordinal()] = year;
            }
            metadata.appendChild(mGenericElement(row, COL.ReleaseYear, false));
        }
        // ReleaseHistoryOriginal ---> ReleaseHistory/Date
        if (!row[COL.ReleaseHistoryOriginal.ordinal()].equals("")) { // optional
            String date = normalizeDate(row[COL.ReleaseHistoryOriginal.ordinal()]);
            if (date == null) {
                throw new ParseException("Invalid ReleaseHistoryOriginal: " + row[COL.ReleaseHistoryOriginal.ordinal()], 0);
            } else {
                row[COL.ReleaseHistoryOriginal.ordinal()] = date;
            }
            Element rh = dom.createElement("ReleaseHistory");
            Element rt = dom.createElement("ReleaseType");
            Text tmp = dom.createTextNode("original");
            rt.appendChild(tmp);
            rh.appendChild(rt);
            rh.appendChild(mGenericElement(row, COL.ReleaseHistoryOriginal, false));
            metadata.appendChild(rh);
        }
        // ReleaseHistoryPhysicalHV ---> ReleaseHistory/Date
        if (!row[COL.ReleaseHistoryPhysicalHV.ordinal()].equals("")) { // optional
            String date = normalizeDate(row[COL.ReleaseHistoryPhysicalHV.ordinal()]);
            if (date == null) {
                throw new ParseException("Invalid ReleaseHistoryPhysicalHV: " + row[COL.ReleaseHistoryPhysicalHV.ordinal()], 0);
            } else {
                row[COL.ReleaseHistoryPhysicalHV.ordinal()] = date;
            }
            Element rh = dom.createElement("ReleaseHistory");
            Element rt = dom.createElement("ReleaseType");
            Text tmp = dom.createTextNode("DVD");
            rt.appendChild(tmp);
            rh.appendChild(rt);
            rh.appendChild(mGenericElement(row, COL.ReleaseHistoryPhysicalHV, false));
            metadata.appendChild(rh);
        }
        // RatingSystem ---> Ratings
        if (row[COL.RatingSystem.ordinal()].equals("")) { // optional
            if (!(row[COL.RatingValue.ordinal()].equals("") && row[COL.RatingReason.ordinal()].equals("")))
                throw new ParseException("RatingSystem not specified", 0);
        } else {
            Element ratings = dom.createElement("Ratings");
            Element rat = dom.createElement("Rating");
            ratings.appendChild(rat);
            Comment comment = dom.createComment("Ratings Region derived from Spreadsheet Territory value");
            rat.appendChild(comment);
            Element region = dom.createElement("Region");
            Text tmp = dom.createTextNode(territory);
            region.appendChild(tmp);
            rat.appendChild(region);
            rat.appendChild(mGenericElement(row, COL.RatingSystem, true));
            rat.appendChild(mGenericElement(row, COL.RatingValue, true));
            Element reason = mGenericElement(row, COL.RatingReason, false);
            if (reason != null)
                rat.appendChild(reason);
            metadata.appendChild(ratings);
        }
        // CaptionIncluded/CaptionException
        if ((e = mCaptionsExemptionReason(row, territory)) != null) {
            if (!territory.equals("US")) {
                Comment comment = dom.createComment("Exemption reason specified for non-US territory");
                metadata.appendChild(comment);
            }
            metadata.appendChild(e);
        }
        // TotalRunTime
        if ((e = mRunLength(row, COL.TotalRunTime)) != null) {
            metadata.appendChild(e);
        }
        // Attach generated Metadata node
        asset.appendChild(metadata);
    } /* mMovieAsset() */

    private Node  mEpisodeAsset(String[] row, Element asset) throws Exception {
        return null;
    }

    private Node mAsset(String[] row) throws Exception {
        Comment comment = null;
        String workType = row[COL.WorkType.ordinal()];

        if (!(workType.equals("Movie") || workType.equals("Episode"))) {
            if (strict) {
                throw new ParseException("invalid workType: " + workType, 0);
            } else {
                Pattern pat = Pattern.compile("^\\s*(movie|episode)\\s*$", Pattern.CASE_INSENSITIVE);
                Matcher m = pat.matcher(workType);
                if (m.matches()) {
                    comment = dom.createComment("corrected from '" + workType + "'");
                    workType = m.group(1).substring(0, 1).toUpperCase() + m.group(1).substring(1).toLowerCase();
                }
                else throw new ParseException("invalid workType: " + workType, 0);
            }
        }
        Element asset = dom.createElement("Asset");
        if (comment != null)
            asset.appendChild(comment);
        Element wt = dom.createElement("WorkType");
        Text tmp = dom.createTextNode(workType);
        wt.appendChild(tmp);
        asset.appendChild(wt);
        if (workType.equals("Movie"))
            mMovieAsset(row, asset);
        else
            mEpisodeAsset(row, asset);
        return asset;
    } /* mAsset() */

    private Node makeTextTerm(String name, String value) {
        Element e = dom.createElement("Term");   
        Attr attr = dom.createAttribute("termName");
        attr.setValue(name);
        e.setAttributeNode(attr);
 
        Element e2 = dom.createElement("Text");
        Text tmp = dom.createTextNode(value);
        e2.appendChild(tmp);
        e.appendChild(e2);
        return e;
    }

    private Node makeDurationTerm(String name, String value) {
        Element e = dom.createElement("Term");   
        Attr attr = dom.createAttribute("termName");
        attr.setValue(name);
        e.setAttributeNode(attr);
 
        // XXX validate
        Element e2 = dom.createElement("Duration");
        Text tmp = dom.createTextNode(value);
        e2.appendChild(tmp);
        e.appendChild(e2);
        return e;
    }

    private Node makeLanguageTerm(String name, String value) {
        Element e = dom.createElement("Term");   
        Attr attr = dom.createAttribute("termName");
        attr.setValue(name);
        e.setAttributeNode(attr);
 
        // XXX validate
        Element e2 = dom.createElement("Language");
        Text tmp = dom.createTextNode(value);
        e2.appendChild(tmp);
        e.appendChild(e2);
        return e;
    }

    private Node makeEventTerm(String name, String dateTime) {
        Element e = dom.createElement("Term");   
        Attr attr = dom.createAttribute("termName");
        attr.setValue(name);
        e.setAttributeNode(attr);
 
        Element e2 = dom.createElement("Event");
        // XXX validate dateTime
        Text tmp = dom.createTextNode(dateTime);
        e2.appendChild(tmp);
        e.appendChild(e2);
        return e;
    }

    private Node makeMoneyTerm(String name, String value, String currency) {
        Element e = dom.createElement("Term");   
        Attr attr = dom.createAttribute("termName");
        attr.setValue(name);
        e.setAttributeNode(attr);
 
        Element e2 = dom.createElement("Money");
        attr = dom.createAttribute("currency");
        attr.setValue(currency);
        e2.setAttributeNode(attr);
        Text tmp = dom.createTextNode(value);
        e2.appendChild(tmp);
        e.appendChild(e2);
        return e;
    }

    private Node mTransaction(String[] row) throws Exception {
        Element transaction = dom.createElement("Transaction");
        Attr attr;
        
        // LicenseType
        String val = row[COL.LicenseType.ordinal()];
        if (!val.matches("^\\s*EST|VOD|SVOD|POEST\\s*$"))
            throw new ParseException("invalid LicenseType: " + val, 0);
        Element e = dom.createElement(COL.LicenseType.toString());
        Text tmp = dom.createTextNode(val);
        e.appendChild(tmp);
        transaction.appendChild(e);

        // Description
        // XXX this can be optional in SS, but not in XML
        val = row[COL.Description.ordinal()];
        if (val.equals(""))
            row[COL.Description.ordinal()] = MISSING;
        e = mGenericElement(row, COL.Description, true);
        transaction.appendChild(e);

        // Territory
        // ISO 3166-1 alpha-2
        transaction.appendChild(mGenericElement(row, COL.Territory, true));

        // Start or StartCondition
        // XXX XML allows empty value
        val = row[COL.Start.ordinal()];
        if (val.equals("")) {
            throw new ParseException("Missing Start date: ", 0);
            // e = dom.createElement(COL.Start.toString());
            // tmp = dom.createTextNode("Immediate");
            // e.appendChild(tmp);
            // transaction.appendChild(e);
        } else {
            String date = normalizeDate(val);
            if (date == null)
                throw new ParseException("Invalid Start date: " + row[COL.ReleaseHistoryOriginal.ordinal()], 0);
            row[COL.ReleaseHistoryOriginal.ordinal()] = date;
            e = dom.createElement(COL.Start.name());  //[sic] yes name
            tmp = dom.createTextNode(date + "T00:00:00");
            e.appendChild(tmp);
            transaction.appendChild(e);
       }

        // End or EndCondition
        val = row[COL.End.ordinal()];
        if (val.equals(""))
            throw new ParseException("End date may not be null", 0);
        String date = normalizeDate(val);
        if (date != null) {
            e = dom.createElement(COL.End.name());  //[sic] yes name
            tmp = dom.createTextNode(date + "T00:00:00");
            e.appendChild(tmp);
            transaction.appendChild(e);
        } else if(val.matches("^\\s*Open|ESTStart|Immediate\\s*$")) {
            e = dom.createElement(COL.End.toString());  //[sic] yes toString
            tmp = dom.createTextNode(val);
            e.appendChild(tmp);
            transaction.appendChild(e);
        } else {
            throw new ParseException("Invalid End Condition " + val, 0);
        }

        // StoreLanguage
        // XXX RFC 1766
        if ((e = mGenericElement(row, COL.StoreLanguage, false)) != null)
            transaction.appendChild(e);

        // LicenseRightsDescription
        val = row[COL.LicenseRightsDescription.ordinal()];
        if (!val.equals("")) {
            if (!LRD.contains(val))
                throw new ParseException("invalid LicenseRightsDescription " + val, 0);
            e = dom.createElement(COL.LicenseRightsDescription.toString());
            tmp = dom.createTextNode(val);
            e.appendChild(tmp);
            transaction.appendChild(e);
        }

        // FormatProfile
        val = row[COL.FormatProfile.ordinal()];
        if (!val.matches("^\\s*SD|HD|3D\\s*$"))
            throw new ParseException("invalid FormatProfile: " + val, 0);
        e = dom.createElement(COL.FormatProfile.toString());
        tmp = dom.createTextNode(val);
        e.appendChild(tmp);
        transaction.appendChild(e);
       
        // ContractID
        if ((e = mGenericElement(row, COL.ContractID, false)) != null)
            transaction.appendChild(e);

        // Term(s)
        // PriceType term
        // XXX validation
        Element e2 = null;
        val = row[COL.PriceType.ordinal()].toLowerCase();
        Pattern pat = Pattern.compile("^\\s*(tier|category|wsp|srp)\\s*$", Pattern.CASE_INSENSITIVE);
        Matcher m = pat.matcher(val);
        String termsType = null, termName = null;
        if (!m.matches())
            throw new ParseException("Invalid PriceType: " + val, 0);
        switch(val) {
        case "tier":
        case "category":
            transaction.appendChild(makeTextTerm(val, row[COL.PriceValue.ordinal()]));
            break;
        case "wsp":
        case "srp":
            transaction.appendChild(makeMoneyTerm(val, row[COL.PriceValue.ordinal()], "USD"));
        }

        // SuppressionLiftDate term
        // XXX validate
        val = row[COL.SuppressionLiftDate.ordinal()];
        if (!val.equals("")) {
            transaction.appendChild(makeEventTerm(COL.SuppressionLiftDate.toString(), normalizeDate(val)));
        }

        // Any Term
        val = row[COL.Any.ordinal()];
        if (!val.equals("")) {
            transaction.appendChild(makeTextTerm(COL.Any.toString(), val));
        }

        // RentalDuration Term
        val = row[COL.RentalDuration.ordinal()];
        if (!val.equals("")) {
            transaction.appendChild(makeDurationTerm(COL.RentalDuration.toString(), val));
        }

        // WatchDuration Term
        val = row[COL.WatchDuration.ordinal()];
        if (!val.equals("")) {
            transaction.appendChild(makeDurationTerm(COL.WatchDuration.toString(), val));
        }

        // HoldbackLanguage Term
        val = row[COL.HoldbackLanguage.ordinal()];
        if (!val.equals("")) {
            transaction.appendChild(makeLanguageTerm(COL.HoldbackLanguage.toString(), val));
        }

        // HoldbackExclusionLanguage Term
        val = row[COL.HoldbackExclusionLanguage.ordinal()];
        if (!val.equals("")) {
            transaction.appendChild(makeLanguageTerm(COL.HoldbackExclusionLanguage.toString(), val));
        }

        // OtherInstructions
        if ((e = mGenericElement(row, COL.OtherInstructions, false)) != null)
            transaction.appendChild(e);

        return transaction;
    }

    //---------------------------------------
    private void availGen(String[] row) throws Exception {
        Element avail = dom.createElement("Avail");
        root.appendChild(avail);
        // ALID
        avail.appendChild(mALID(row[COL.AvailID.ordinal()]));
        // Disposition
        avail.appendChild(mDisposition(row[COL.EntryType.ordinal()]));
        // Licensor
        avail.appendChild(mLicensor(row[COL.DisplayName.ordinal()]));
        // Exception Flag
        Node n = mException(row[COL.ExceptionFlag.ordinal()]);
        if (n != null)
            avail.appendChild(n);
        // Asset
        n = mAsset(row);
        if (n != null)
            avail.appendChild(n);
        n = mTransaction(row);
        if (n != null)
            avail.appendChild(n);

    }

    public void makeXML(ArrayList<Object> rows) throws Exception {
		
    	//get an instance of factory
        DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
        dbf.setValidating(true);
        try {
            //get an instance of builder
            DocumentBuilder db = dbf.newDocumentBuilder();

            //create an instance of DOM
            dom = db.newDocument();
            root = dom.createElement("AvailList");
            // Make it the root element of this new document
            Attr xmlns = dom.createAttribute("xmlns");
            xmlns.setValue("http://www.movielabs.com/schema/avails/v2.0/avails");
            root.setAttributeNode(xmlns);

            xmlns = dom.createAttribute("xmlns:xsi");
            xmlns.setValue("http://www.w3.org/2001/XMLSchema-instance");
            root.setAttributeNode(xmlns);

            xmlns = dom.createAttribute("xmlns:md");
            xmlns.setValue("http://www.movielabs.com/schema/md/v2.3/md");
            root.setAttributeNode(xmlns);

            xmlns = dom.createAttribute("xmlns:mdmec");
            xmlns.setValue("http://www.movielabs.com/schema/mdmec/v2.3");
            root.setAttributeNode(xmlns);

            xmlns = dom.createAttribute("xsi:schemaLocation");
            xmlns.setValue("http://www.movielabs.com/schema/avails/v2.0/avails " + 
                           "http://www.movielabs.com/schema/avails/v2.0/avails-v2.0.xsd");
            root.setAttributeNode(xmlns);

            dom.appendChild(root);

            for (Object r : rows) {
                String[] row = (String[]) r;
                availGen(row);
            }

            // Use a Transformer for output
            TransformerFactory tFactory = TransformerFactory.newInstance();
            Transformer transformer = tFactory.newTransformer();
            transformer.setOutputProperty(OutputKeys.INDENT, "yes");
            transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "4");

            DOMSource source = new DOMSource(dom);
            //StreamResult result = new StreamResult(System.out);
            StreamResult result = new StreamResult(new FileOutputStream("testout.xml"));
            transformer.transform(source, result);

        } catch(ParserConfigurationException pce) {
            //dump it
            System.out.println("Error while trying to instantiate DocumentBuilder " + pce);
            System.exit(1);
        } catch (TransformerConfigurationException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (TransformerException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
}
