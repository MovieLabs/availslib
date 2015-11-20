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

public class AvailsSheet {
    private ArrayList<SheetRow> rows;
    private AvailSS parent;
    private String name;

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

    public AvailsSheet(AvailSS parent, String name) {
        this.parent = parent;
        this.name = name;
        rows = new ArrayList<SheetRow>();
    }

    public AvailsSheet(AvailSS parent, String name, int initialRows) {
        this.parent = parent;
        this.name = name;
        rows = new ArrayList<SheetRow>(initialRows);
    }

    public boolean isAvail(String[] fields) {
        String t = fields[COL.Territory.ordinal()];
        boolean ret = (t != null) && 
            !(t.equals("AvailTrans") || t.equals("Territory") || t.substring(0, 2).equals("//"));
        return ret;
    }

    public String getName() {
        return name;
    }

    public ArrayList<SheetRow> getRows() {
        return rows;
    }

    private void log(String s, boolean bail) throws Exception {
        s = String.format("Sheet %s: %s", name, s);
        parent.getLogger().warn(s);
        if (bail)
            throw new ParseException(s, 0);
    }

    public void addRow(String[] fields) throws Exception {

        String workType = fields[COL.WorkType.ordinal()];

        if (!(workType.equals("Movie") || workType.equals("Episode"))) {
            if (parent.getCleanupData()) {
                Pattern pat = Pattern.compile("^\\s*(movie|episode|season)\\s*$", Pattern.CASE_INSENSITIVE);
                Matcher m = pat.matcher(workType);
                if (m.matches()) {
                    log("corrected from '" + workType + "'", false);
                    workType = m.group(1).substring(0, 1).toUpperCase() + m.group(1).substring(1).toLowerCase();
                } else {
                    log("invalid workType: " + workType, true);
                    return;
                }
            } else {
                log("invalid workType: " + workType, true);
                return;
            }
        }
        SheetRow sr;
        switch(workType) {
        case "Movie": 
            sr = new Movie(this, "Movie", rows.size() + 1, fields);
            break;
        case "Episode":
            sr = new Episode(this, "Episode", rows.size() + 1, fields);
            break;
        case "Season":
        	sr = new Season(this, "Season", rows.size() + 1, fields);
        	break;
        default:
            log("invalid workType: " + workType, true);
            return;
        }
        rows.add(sr);
    }

    public AvailSS getAvailSS() {
        return parent;
    }

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
