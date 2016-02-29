/*
 * Copyright (c) 2015 MovieLabs
 * 
 * Permission is hereby granted, free of charge, to any person obtaining
 * a copy of this software and associated documentation files (the
 * "Software"), to deal in the Software without restriction, including
 * without limitation the rights to use, copy, modify, merge, publish,
 * distribute, sublicense, and/or sell copies of the Software, and to
 * permit persons to whom the Software is furnished to do so, subject to
 * the following conditions:
 * 
 * The above copyright notice and this permission notice shall be
 * included in all copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
 * EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
 * MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
 * NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
 * LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
 * OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
 * WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 *
 * Author: Paul Jensen <pgj@movielabs.com>
 */

package com.movielabs.avails;
import java.util.*;
import java.io.*;
import java.text.ParseException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.*;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.w3c.dom.*;

/**
 * Represents an individual sheet of an Excel spreadsheet
 */
public class AvailsSheet {
    private ArrayList<SheetRow> rows;
    private AvailSS parent;
    private String name;

    /**
     * An enum used to represent offsets into an array of spreadsheet cells.  The name() method returns
     * the Column name (based on Avails spreadsheet representation; the toString() method returns a
     * corresponding or related XML element; and ordinal() returns the offset
     */
    private enum COL {
        DisplayName ("DisplayName"),                                   //  0
        StoreLanguage ("StoreLanguage"),                               //  1
        Territory ("md:country"),                                      //  2
        WorkType ("WorkType"),                                         //  3
        EntryType ("EntryType");                                       //  4


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
    } /* COL */

    /**
     * Create an object representing a single sheet of a spreadsheet
     * @param parent the parent Spreadsheet object
     * @param name the name of the spreadsheet
     */
    public AvailsSheet(AvailSS parent, String name) {
        this.parent = parent;
        this.name = name;
        rows = new ArrayList<SheetRow>();
    }

    /**
     * Create an object representing a single sheet of a spreadsheet
     * @param parent the parent Spreadsheet object
     * @param name the name of the spreadsheet
     * @param initialRows an estimate of the number of rows in this sheet
     */
    public AvailsSheet(AvailSS parent, String name, int initialRows) {
        this.parent = parent;
        this.name = name;
        rows = new ArrayList<SheetRow>(initialRows);
    }

    /**
     * Determine if a spreadsheet row contains an avail
     * @param fields an array containing the raw values of a spreadsheet row
     * @return true iff the row is an avail based on the contents of the Territory column
     */
    public boolean isAvail(String[] fields) {
        String t = fields[COL.Territory.ordinal()];
        boolean ret = (t != null) && 
            !(t.equals("AvailTrans") || t.equals("Territory") || 
              (t.length() >= 2 && t.substring(0, 2).equals("//")));
        return ret;
    }

    public String getName() {
        return name;
    }

    /**
     * Get a an array of objects representing each row of this sheet
     * @return an array containing all the SheetRow objects in this sheet
     */
    public ArrayList<SheetRow> getRows() {
        return rows;
    }

    /**
     * helper routine to create a log entry
     * @param s the data to be logged
     * @param bail if true, throw a ParseException after logging the message
     * @throws ParseException if bail is true
     */
    private void log(String s, int rowNum, boolean bail) throws Exception {
        s = String.format("Sheet %s Row %5d: %s", name, rowNum, s);
        parent.getLogger().warn(s);
        if (bail)
            throw new ParseException(s, 0);
    }

    /**
     * Add a row of spreadsheet data
     * @param fields an array containing the raw values of a spreadsheet row 
     * @param rowNum the row number from the source spreadsheet
     * @throws Exception if an invalid workType is encountered
     *         (currently, only 'movie', 'episode', or 'season' are accepted)
     */
    public void addRow(String[] fields, int rowNum) throws Exception {

        String workType = fields[COL.WorkType.ordinal()];

        if (!(workType.equals("Movie") || workType.equals("Episode") || workType.equals("Season"))) {
            if (parent.getCleanupData()) {
                Pattern pat = Pattern.compile("^\\s*(movie|episode|season)\\s*$", Pattern.CASE_INSENSITIVE);
                Matcher m = pat.matcher(workType);
                if (m.matches()) {
                    log("corrected from '" + workType + "'", rowNum, false);
                    workType = m.group(1).substring(0, 1).toUpperCase() + m.group(1).substring(1).toLowerCase();
                } else {
                    log("invalid workType: '" + workType + "'", rowNum, false);
                    return;
                }
            } else {
                log("invalid workType: '" + workType + "'", rowNum, false);
                return;
            }
        }
        SheetRow sr;
        switch(workType) {
        case "Movie": 
            sr = new Movie(this, "Movie", rowNum, fields);
            break;
        case "Episode":
            sr = new Episode(this, "Episode", rowNum, fields);
            break;
        case "Season":
            sr = new Season(this, "Season", rowNum, fields);
            break;
        default:
            log("invalid workType: " + workType, rowNum, true);
            return;
        }
        rows.add(sr);
    }

    /**
     * get the parent Spreadsheet object for this sheet
     * @return the parent of this sheet
     */
    public AvailSS getAvailSS() {
        return parent;
    }

    /**
     * Create an Avails XML document based on the data in this spreadsheet
     * @param shortDesc a short description that will appear in the document
     * @return a JAXP document
     * @throws Exception if any errors are encountered
     */
    public Document makeXML(String shortDesc) throws Exception {
        Document dom = null;
        Element root;

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

            int lineNo = 0;
            for (SheetRow r : rows) {
                lineNo++;
                if (shortDesc != null & !shortDesc.equals(""))
                    r.setShortDesc(shortDesc);
                Element e = r.makeAvail(dom);
                root.appendChild(e);
            }
        } catch(ParserConfigurationException pce) {
            //dump it
            System.out.println("Error while trying to instantiate DocumentBuilder " + pce);
            System.exit(1);
        }

        return dom;
    }

    /**
     * Create an Avails XML file based on the data in this spreadsheet
     * @param xmlFile the name of the created XML output file
     * @param shortDesc a short description that will appear in the document
     * @throws Exception if any errors are encountered
     */
    public void makeXMLFile(String xmlFile, String shortDesc) throws Exception {
        try {
            Document dom = makeXML(shortDesc);
            
            // Use a Transformer for output
            TransformerFactory tFactory = TransformerFactory.newInstance();
            Transformer transformer = tFactory.newTransformer();
            transformer.setOutputProperty(OutputKeys.INDENT, "yes");
            transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "4");
            
            DOMSource source = new DOMSource(dom);
            //StreamResult result = new StreamResult(System.out);
            StreamResult result = new StreamResult(new FileOutputStream(xmlFile));
            transformer.transform(source, result);

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
