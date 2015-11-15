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
import org.apache.logging.log4j.*;

public abstract class SheetRow {
    protected AvailsSheet parent;
    protected int lineNo;
    protected String[] fields;
    protected Logger log;
    protected Document dom;
    protected Element root;
    protected static final String MISSING = "---missing---";
    protected boolean exitOnError;
    protected boolean cleanupData;
    protected String workType;
    protected String shortDesc;
    
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

    public SheetRow(AvailsSheet parent, String workType, int lineNo, String[] fields) {
        this.parent = parent;
        this.workType = workType;
        this.lineNo = lineNo;
        this.fields = fields;
        this.log = getLogger();
        this.exitOnError = parent.getAvailSS().getExitOnError();
        this.cleanupData = parent.getAvailSS().getCleanupData();
        shortDesc = ""; // default
    }

    public String getWorkType() {
        return workType;
    }

    public Logger getLogger() {
        return parent.getAvailSS().getLogger();
    }

    public void setShortDesc(String shortDesc) {
        this.shortDesc = shortDesc;
    }

    public String getShortDesc(String shortDesc) {
        return shortDesc;
    }

    /**
     * Create an XML element
     * @param name the name of the element
     * @param val the value of the element
     * @param mandatory if true, indicates this is a required field,
     *        and if it is null an error will be reported
     * @return the created element, or null if there is an error
     * @throws ParseException if there is an error and abort-on-error policy is in effect
     */
    protected Element mGenericElement(String name, String val, boolean mandatory) throws Exception {
        if (val.equals("")) {
            if (mandatory)
                reportError("missing required value on element: " + name);
            else
                return null;
        }
        Element tia = dom.createElement(name);
        Text tmp = dom.createTextNode(val);
        tia.appendChild(tmp);
        return tia;
    } 

    /* **************************************
     * Node-generating methods
     ****************************************/

    /**
     * Creates an Avails/ALID XML node
     * @param availID the value of the ALID
     * @return the generated XML element
     */
    protected Element mALID(String availID) {
        if (availID.equals(""))
            availID = MISSING;
        Text tmp = dom.createTextNode(availID);
        Element ALID = dom.createElement("ALID");
        ALID.appendChild(tmp);
        return ALID;
    }

    /**
     * Creates an Avails/Disposition XML node
     * @param entryType string (controlled vocabulary) indicating whether this Avail is new, and update,
     *        or a deletion
     * @return the created XML node
     * @throws ParseException if there is an error and abort-on-error policy is in effect
     */
    protected Element mDisposition(String entryType) throws Exception {
        Comment comment = null;
        boolean err = false;

        if (!(entryType.equals("Full Extract") || entryType.equals("Full Delete"))) {
            if (cleanupData) {
                Pattern pat = Pattern.compile("^\\s*full\\s+(extract|delete)\\s*$", Pattern.CASE_INSENSITIVE);
                Matcher m = pat.matcher(entryType);
                if (m.matches()) {
                    comment = dom.createComment("corrected from '" + entryType + "'");
                    if (m.group(1).equalsIgnoreCase("extract"))
                        entryType = "Full Extract";
                    else if (m.group(1).equalsIgnoreCase("delete"))
                        entryType = "Full Delete";
                    else
                        err = true;
                } else {
                    err = true;
                }
            } else {
                err = true;
            }
        }
        if (err)
            reportError("invalid Disposition: " + entryType);
        Element disp = dom.createElement("Disposition");
        Element entry = dom.createElement("EntryType");
        Text tmp = dom.createTextNode(entryType);
        entry.appendChild(tmp);
        disp.appendChild(entry);
        if (comment != null)
            disp.appendChild(comment);
        return disp;
    }

    /**
     * Create an Avails Licensor XML element with a md:DisplayName element child, and populate thei latter 
     * with the DisplayName 
     * @param displayName the name to be held in the DisplayName child node of Licensor
     * @return the created Licensor element
     * @throws ParseException if there is an error and abort-on-error policy is in effect
     */
    protected Element mPublisher(String name, String displayName, boolean mandatory) throws Exception {
        if (displayName.equals("")) {
            if (mandatory)
                reportError("missing md:DisplayName");
            else
                return null;
        }
        Element licensor = dom.createElement(name);
        Element e = dom.createElement("md:DisplayName");
        Text tmp = dom.createTextNode(displayName);
        e.appendChild(tmp);
        licensor.appendChild(e);
        // XXX ContactInfo mandatory but can't get this info from the spreadsheet
        e = dom.createElement("mdmec:ContactInfo");
        Element e2 = dom.createElement("md:Name");
        e.appendChild(e2);
        e2 = dom.createElement("md:PrimaryEmail");
        e.appendChild(e2);
        licensor.appendChild(e);

        return licensor;
    }

    protected abstract Element mAssetBody(Element asset) throws Exception;

    protected Element mAssetHeader() throws Exception {
        Element asset = dom.createElement("Asset");
        Element wt = dom.createElement("WorkType");
        Text tmp = dom.createTextNode(workType);
        wt.appendChild(tmp);
        asset.appendChild(wt);
        return mAssetBody(asset);
    } /* mAsset() */

    protected abstract Element mTransactionBody(Element transaction) throws Exception;

    protected Element mTransactionHeader() throws Exception {
        Element transaction = dom.createElement("Transaction");
        return mTransactionBody(transaction);
    } /* mTransaction() */

    // ------------------------ Transaction-related methods

    protected Element mLicenseType(String val) throws Exception {
        if (!val.matches("^\\s*EST|VOD|SVOD|POEST\\s*$"))
            reportError("invalid LicenseType: " + val);
        Element e = dom.createElement("LicenseType");
        Text tmp = dom.createTextNode(val);
        e.appendChild(tmp);
        return e;
    }

    // XXX ISO 3166-1 alpha-2
    protected Element mTerritory(String val) throws Exception {
        Element e = dom.createElement("Territory");
        Element e2 = mGenericElement("md:country", val, true);
        e.appendChild(e2);
        return e;
    }

    protected Element mDescription(String val) throws Exception {
        // XXX this can be optional in SS, but not in XML
        if (val.equals(""))
            val = MISSING;
        return mGenericElement("Description", val, true);
    }

    // XXX XML allows empty value
    // XXX cleanupData code not added
    /**
     * Generate Start or StartCondition
     * @param val start time
     * @return a Start element
     */
    protected Element mStart(String val) throws Exception {
        Element e = null;
        if (val.equals("")) {
            reportError("Missing Start date: ");
            // e = dom.createElement("StartCondition");
            // tmp = dom.createTextNode("Immediate");
        } else {
            String date = normalizeDate(val);
            if (date == null)
                reportError("Invalid Start date: " + val);
            e = dom.createElement("Start");  //[sic] yes name
            Text tmp = dom.createTextNode(date + "T00:00:00");
            e.appendChild(tmp);
        }
        return e;
    }

    // XXX cleanupData code not added
    protected Element mEnd(String val) throws Exception {
        Element e = null;
        Text tmp;
        if (val.equals(""))
            reportError("End date may not be null");

        String date = normalizeDate(val);
        if (date != null) {
            e = dom.createElement("End");
            tmp = dom.createTextNode(date + "T00:00:00");
            e.appendChild(tmp);
        } else if (val.matches("^\\s*Open|ESTStart|Immediate\\s*$")) {
            e = dom.createElement("EndCondition");
            tmp = dom.createTextNode(val);
            e.appendChild(tmp);
        } else {
            reportError("Invalid End Condition " + val);
        }
        return e;
    }

    // XXX RFC 1766 validation needed
    protected Element mStoreLanguage(String val) throws Exception {
        return mGenericElement("StoreLanguage", val, false);
    }

    // XXX cleanupData code not added
    protected Element mLicenseRightsDescription(String val) throws Exception {
        if (!LRD.contains(val))
            reportError("invalid LicenseRightsDescription " + val);
        Element e = dom.createElement("LicenseRightsDescription");
        Text tmp = dom.createTextNode(val);
        e.appendChild(tmp);
        return e;
    }

    // XXX cleanupData code not added
    protected Element mFormatProfile(String val) throws Exception {
        if (!val.matches("^\\s*SD|HD|3D\\s*$"))
            reportError("invalid FormatProfile: " + val);
        Element e = dom.createElement("FormatProfile");
        Text tmp = dom.createTextNode(val);
        e.appendChild(tmp);
        return e;
    }

    // XXX cleanupData code not added
    protected Element mPriceType(String priceType, String priceVal) throws Exception {
        Element e = null;
        priceType = priceType.toLowerCase();
        Pattern pat = Pattern.compile("^\\s*(tier|category|wsp|srp)\\s*$", Pattern.CASE_INSENSITIVE);
        Matcher m = pat.matcher(priceType);
        if (!m.matches())
            reportError("Invalid PriceType: " + priceType);
        switch(priceType) {
        case "tier":
        case "category":
            e = makeTextTerm(priceType, priceVal);
            break;
        case "wsp":
        case "srp":
            // XXX set currency properly
            e = makeMoneyTerm(priceType, priceVal, "USD");
        }
        return e;
    }

    protected Element makeTextTerm(String name, String value) {
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

    protected Element makeDurationTerm(String name, String value) throws Exception {
        int hours;
        try {
            hours = normalizeInt(value);
        } catch(NumberFormatException e) {
            reportError(" invalid duration: " + value);
            return null;
        }
        value = String.format("PT%dH", hours);
            
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

    protected Element makeLanguageTerm(String name, String value) {
        // XXX validate
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

    protected Element makeEventTerm(String name, String dateTime) {
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

    protected Element makeMoneyTerm(String name, String value, String currency) {
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

    // ------------------

    protected Element mRunLength(String val) throws Exception {
        Element ret = dom.createElement("RunLength");
        Text tmp = null;
        if (val.equals("")) { // XXX ugly hack
            tmp = dom.createTextNode("PT0H");
        } else {
            Pattern dur = Pattern.compile("^\\s*(\\d{1,2}):(\\d{1,2}):(\\d{1,2})\\s*$");
            Matcher m = dur.matcher(val);
            if (m.matches()) {
                int hour, min, sec;
                try {
                    hour = normalizeInt(m.group(1));
                    min = normalizeInt(m.group(2));
                    sec = normalizeInt(m.group(3));
                } catch(NumberFormatException s) {
                    hour = min = sec = 60;
                }
                if (min > 59 || sec > 59)
                    reportError("invalid duration string " + val);
                String d =  String.format("PT%dH%dM%dS", hour, min, sec);
                tmp = dom.createTextNode(d);
            } else {
                reportError("invalid duration string (verify Excel cell format is 'Text'): " + val);
                ret = null;
            }
        }
        ret.appendChild(tmp);
        return ret;
    }

    protected Element mLocalizationType(Element parent, String loc) throws Exception {
        if (loc.equals(""))
            return null;
        if (!(loc.equals("sub") || loc.equals("dub") || loc.equals("subdub") || loc.equals("any"))) {
            if (cleanupData) {
                Pattern pat = Pattern.compile("^\\s*(sub|dub|subdub|any)\\s*$", Pattern.CASE_INSENSITIVE);
                Matcher m = pat.matcher(loc);
                if (m.matches()) {
                    Comment comment = dom.createComment("corrected from '" + loc + "'");
                    loc = m.group(1).toLowerCase();
                    parent.appendChild(comment);
                }
            } else {
                reportError("invalid LocalizationOffering value: " + loc);
            }
        }
        return mGenericElement("LocalizationOffering", loc, false);
    }

    protected Element mCaptionsExemptionReason(String captionIncluded, String captionExemption,
                                               String territory) throws Exception {
        // XXX clarify policy regarding non-US captioning
        //        if (territory.equals("US")) {  // check captions
            switch(yesorno(captionIncluded)) {
            case 1:
                if (!captionExemption.equals(""))
                    reportError("CaptionExemption specified without CaptionIncluded");
                break;
            case 0:
                if (captionExemption.equals(""))
                    reportError("CaptionExemption not specified");
                int exemption;
                try {
                    exemption = normalizeInt(captionExemption);
                } catch(NumberFormatException s) {
                    exemption = -1;
                }
                if (exemption < 1 || exemption > 6)
                    reportError("Invalid CaptionExamption Value: " + exemption);
                Element capex = dom.createElement("USACaptionsExemptionReason");
                Text tmp = dom.createTextNode(Integer.toString(exemption));
                capex.appendChild(tmp);
                return capex;
            default:
                reportError("CaptionExemption specified without CaptionIncluded");
            }
            // } else {
            // if (captionIncluded.equals("") && captionExemption.equals(""))
            //     return null;
            // else
            //     reportError("CaptionIncluded/Exemption should only be specified in US");
            // }
            return null;
    }


    /**
     * Create an Avails ExceptionFlag element
     * @param exceptionFlag a string indicating whether the flag is to be set.  It should be "Yes" or "No",
     *        but several case and whitespace variants will be tolerated.
     * @return the created ExceptionFlag element, or null if there is an error
     * @throws ParseException if the supplied value can't be mapped to
     * a boolean and abort-on-error policy is in effect
     */
    protected Element mExceptionFlag(String exceptionFlag) throws Exception {
        switch(yesorno(exceptionFlag)) {
        case 0:
            return null;
        case 1:
            Element eFlag = dom.createElement("ExceptionFlag");
            Text tmp = dom.createTextNode("true");
            eFlag.appendChild(tmp);
            return eFlag;
        default:
            reportError("invalid ExceptionFlag");
            return null;
        }
    }


    /* **************************************
     * Helper methods
     ****************************************/

    /**
     * logs an error and potentially throws an exception.  The error
     * is decorated with the current Sheet name and the row being
     * processed
     * @param s the error message
     * @throws ParseException if exit-on-error policy is in effect
     */
    protected void reportError(String s) throws Exception {
        s = String.format("Line %05d: %s", lineNo, s);
        log.warn(s);
        if (exitOnError)
            throw new ParseException(s, 0);
    }
    

    /**
     * Parse an input string to determine whether "Yes" or "No" is
     * intended.  A variety of case mismatches and leading/trailing
     * whitespace are accepted.
     * @param s the string to be tested
     * @return 1 if "yes" was intended; 0 if "no" was intended; -1 for an invalid pattern.
     */
    protected int yesorno(String s) {
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

    /**
     * Verify that a string represents a valid EIDR, and return
     * compact representation if so.  Leading and trailing whitespace
     * is tolerated, as is the presence/absence of the EIDR "10.5240/"
     * prefix
     * @param s the input string to be tested
     * @return a short EIDR corresponding to the asset if valid; null
     * if not a proper EIDR exception
     */
    protected String normalizeEIDR(String s) {
        Pattern eidr = Pattern.compile("^\\s*(?:10\\.5240/)?((?:(?:\\p{XDigit}){4}-){5}\\p{XDigit})\\s*$");
        Matcher m = eidr.matcher(s);
        if (m.matches()) {
            return m.group(1).toUpperCase();
        } else {
            return null;
        }
    }

    /**
     * Verify that a string represents a valid 4-digit year, and
     * return a canonical representation if so.  Leading and trailing
     * whitespace is tolerated.
     * @param s the input string to be tested
     * @return a 4-digit year value represented as a string
     */
    protected String normalizeYear(String s) {
        Pattern eidr = Pattern.compile("^\\s*(\\d{4})(?:\\.0)?\\s*$");
        Matcher m = eidr.matcher(s);
        if (m.matches()) {
            return m.group(1);
        } else {
            return null;
        }
    }

    protected int normalizeInt(String s) throws Exception {
        Pattern eidr = Pattern.compile("^\\s*(\\d+)(?:\\.0)?\\s*$");
        Matcher m = eidr.matcher(s);
        if (m.matches()) {
            return Integer.parseInt(m.group(1));
        } else {
            throw new NumberFormatException(s);
        }
    }

    protected String normalizeDate(String s) {
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

    protected abstract Element makeAvail(Document dom) throws Exception;
}
