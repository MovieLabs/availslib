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

public class Movie extends SheetRow {

    private enum COL {
        DisplayName ("DisplayName"),                               //  0
        StoreLanguage ("StoreLanguage"),                           //  1
        Territory ("md:country"),                                  //  2
        WorkType ("WorkType"),                                     //  3
        EntryType ("EntryType"),                                   //  4
        TitleInternalAlias ("TitleInternalAlias"),                 //  5
        TitleDisplayUnlimited ("TitleDisplayUnlimited"),           //  6
        LocalizationType ("LocalizationOffering"),                 //  7
        LicenseType ("LicenseType"),                               //  8
        LicenseRightsDescription ("LicenseRightsDescription"),     //  9
        FormatProfile ("FormatProfile"),                           // 10
        Start ("StartCondition"),                                  // 11
        End ("EndCondition"),                                      // 12
        PriceType ("PriceType"),                                   // 13
        PriceValue ("PriceValue"),                                 // 14
        SRP ("SRP"),                                               // 15
        Description ("Description"),                               // 16
        OtherTerms ("OtherTerms"),                                 // 17
        OtherInstructions ("OtherInstructions"),                   // 18
        ContentID ("contentID"),                                   // 19
        ProductID ("EditEIDR-S"),                                  // 20
        EncodeID ("EncodeID"),                                     // 21
        AvailID ("AvailID"),                                       // 22
        Metadata ("Metadata"),                                     // 23
        AltID ("md:Identifier"),                                   // 24
        SuppressionLiftDate ("AnnounceDate"),                      // 25
        SpecialPreOrderFulfillDate ("SpecialPreOrderFulfillDate"), // 26
        ReleaseYear ("ReleaseDate"),                               // 27
        ReleaseHistoryOriginal ("md:Date"),                        // 28
        ReleaseHistoryPhysicalHV ("md:Date"),                      // 29
        ExceptionFlag ("ExceptionFlag"),                           // 30
        RatingSystem ("md:System"),                                // 31
        RatingValue ("md:Value"),                                  // 32
        RatingReason ("md:RatingReason"),                          // 33
        RentalDuration ("RentalDuration"),                         // 34
        WatchDuration ("WatchDuration"),                           // 35
        CaptionIncluded ("USACaptionsExemptionReason"),            // 36
        CaptionExemption ("USACaptionsExemptionReason"),           // 37
        Any ("Any"),                                               // 38
        ContractID ("ContractID"),                                 // 39
        ServiceProvider ("ServiceProvider"),                       // 40
        TotalRunTime ("RunLength"),                                // 41
        HoldbackLanguage ("HoldbackLanguage"),                     // 42
        HoldbackExclusionLanguage ("HoldbackExclusionLanguage");   // 43


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

    protected Element mAssetBody(Element asset) throws Exception{
        Element e;
        String contentID = fields[COL.ContentID.ordinal()];
        if (contentID.equals(""))
            contentID = MISSING;
        Attr attr = dom.createAttribute(COL.ContentID.toString());
        attr.setValue(contentID);
        asset.setAttributeNode(attr);
        Element metadata = dom.createElement("Metadata");
        String territory = fields[COL.Territory.ordinal()];

        // TitleDisplayUnlimited
        // XXX Optional in SS, Required in XML; workaround by assigning it internal alias value
        if (fields[COL.TitleDisplayUnlimited.ordinal()].equals(""))
            fields[COL.TitleDisplayUnlimited.ordinal()] = fields[COL.TitleInternalAlias.ordinal()];
        if ((e = mGenericElement(COL.TitleDisplayUnlimited.toString(), 
                                 fields[COL.TitleDisplayUnlimited.ordinal()], false)) != null)
            metadata.appendChild(e);
        // TitleInternalAlias
        metadata.appendChild(mGenericElement(COL.TitleInternalAlias.toString(),
                                             fields[COL.TitleInternalAlias.ordinal()], true));
        // ProductID --> EditEIDR-S
        if (!fields[COL.ProductID.ordinal()].equals("")) { // optional field
            String productID = normalizeEIDR(fields[COL.ProductID.ordinal()]);
            if (productID == null) {
                reportError("Invalid ProductID: " + fields[COL.ProductID.ordinal()]);
            } else {
                fields[COL.ProductID.ordinal()] = productID;
            }
            metadata.appendChild(mGenericElement(COL.ProductID.toString(),
                                                 fields[COL.ProductID.ordinal()], false));
        }
        // AltID --> AltIdentifier
        if (!fields[COL.AltID.ordinal()].equals("")) { // optional
            Element altID = dom.createElement("AltIdentifier");
            Element cid = dom.createElement("md:Namespace");
            Text tmp = dom.createTextNode(MISSING);
            cid.appendChild(tmp);
            altID.appendChild(cid);
            altID.appendChild(mGenericElement(COL.AltID.toString(),
                                              fields[COL.AltID.ordinal()], false));
            Element loc = dom.createElement("md:Location");
            tmp = dom.createTextNode(MISSING);
            loc.appendChild(tmp);
            altID.appendChild(loc);
            metadata.appendChild(altID);
        }
        // ReleaseYear ---> ReleaseDate
        if (!fields[COL.ReleaseYear.ordinal()].equals("")) { // optional
            String year = normalizeYear(fields[COL.ReleaseYear.ordinal()]);
            if (year == null) {
                reportError("Invalid ReleaseYear: " + fields[COL.ReleaseYear.ordinal()]);
            } else {
                fields[COL.ReleaseYear.ordinal()] = year;
            }
            metadata.appendChild(mGenericElement(COL.ReleaseYear.toString(), 
                                                 fields[COL.ReleaseYear.ordinal()], false));
        }
        // TotalRunTime --> RunLength
        // XXX TotalRunTime is optional, RunLength is required
        if ((e = mRunLength(fields[COL.TotalRunTime.ordinal()])) != null) {
            metadata.appendChild(e);
        }
        // ReleaseHistoryOriginal ---> ReleaseHistory/Date
        if (!fields[COL.ReleaseHistoryOriginal.ordinal()].equals("")) { // optional
            String date = normalizeDate(fields[COL.ReleaseHistoryOriginal.ordinal()]);
            if (date == null) {
                reportError("Invalid ReleaseHistoryOriginal: " + fields[COL.ReleaseHistoryOriginal.ordinal()]);
            } else {
                fields[COL.ReleaseHistoryOriginal.ordinal()] = date;
            }
            Element rh = dom.createElement("ReleaseHistory");
            Element rt = dom.createElement("md:ReleaseType");
            Text tmp = dom.createTextNode("original");
            rt.appendChild(tmp);
            rh.appendChild(rt);
            rh.appendChild(mGenericElement(COL.ReleaseHistoryOriginal.toString(),
                                           fields[COL.ReleaseHistoryOriginal.ordinal()], false));
            metadata.appendChild(rh);
        }
        // ReleaseHistoryPhysicalHV ---> ReleaseHistory/Date
        if (!fields[COL.ReleaseHistoryPhysicalHV.ordinal()].equals("")) { // optional
            String date = normalizeDate(fields[COL.ReleaseHistoryPhysicalHV.ordinal()]);
            if (date == null) {
                reportError("Invalid ReleaseHistoryPhysicalHV: " + fields[COL.ReleaseHistoryPhysicalHV.ordinal()]);
            } else {
                fields[COL.ReleaseHistoryPhysicalHV.ordinal()] = date;
            }
            Element rh = dom.createElement("ReleaseHistory");
            Element rt = dom.createElement("md:ReleaseType");
            Text tmp = dom.createTextNode("DVD");
            rt.appendChild(tmp);
            rh.appendChild(rt);
            rh.appendChild(mGenericElement(COL.ReleaseHistoryPhysicalHV.toString(),
                                           fields[COL.ReleaseHistoryPhysicalHV.ordinal()], false));
            metadata.appendChild(rh);
        }
        // CaptionIncluded/CaptionException
        if ((e = mCaptionsExemptionReason(fields[COL.CaptionIncluded.ordinal()],
                                          fields[COL.CaptionExemption.ordinal()],
                                          territory)) != null) {
            if (!territory.equals("US")) {
                Comment comment = dom.createComment("Exemption reason specified for non-US territory");
                metadata.appendChild(comment);
            }
            metadata.appendChild(e);
        }
        // RatingSystem ---> Ratings
        if (fields[COL.RatingSystem.ordinal()].equals("")) { // optional
            if (!(fields[COL.RatingValue.ordinal()].equals("") && fields[COL.RatingReason.ordinal()].equals("")))
                reportError("RatingSystem not specified");
        } else {
            Element ratings = dom.createElement("Ratings");
            Element rat = dom.createElement("md:Rating");
            ratings.appendChild(rat);
            Comment comment = dom.createComment("Ratings Region derived from Spreadsheet Territory value");
            rat.appendChild(comment);
            Element region = dom.createElement("md:Region");
            // XXX validate country
            Element country = dom.createElement("md:country");
            region.appendChild(country);
            Text tmp = dom.createTextNode(territory);
            country.appendChild(tmp);
            rat.appendChild(region);
            rat.appendChild(mGenericElement(COL.RatingSystem.toString(), 
                                            fields[COL.RatingSystem.ordinal()], true));
            rat.appendChild(mGenericElement(COL.RatingValue.toString(),
                                            fields[COL.RatingValue.ordinal()], true));
            Element reason = mGenericElement(COL.RatingReason.toString(),
                                             fields[COL.RatingReason.ordinal()], false);
            if (reason != null)
                rat.appendChild(reason);
            metadata.appendChild(ratings);
        }
        // EncodeID --> EditEIDR-S
        if (!fields[COL.EncodeID.ordinal()].equals("")) { // optional field
            String encodeID = normalizeEIDR(fields[COL.EncodeID.ordinal()]);
            if (encodeID == null) {
                reportError("Invalid EncodeID: " + fields[COL.EncodeID.ordinal()]);
            } else {
                fields[COL.EncodeID.ordinal()] = encodeID;
            }
            metadata.appendChild(mGenericElement(COL.EncodeID.toString(),
                                                 fields[COL.EncodeID.ordinal()], false));
        }
        // LocalizationType --> LocalizationOffering
        if ((e = mLocalizationType(metadata, fields[COL.LocalizationType.ordinal()])) != null) {
            metadata.appendChild(e);
        }
        // Attach generated Metadata node
        asset.appendChild(metadata);
        return asset;
    } /* mAssetBody() */

    protected Element mTransactionBody(Element transaction) throws Exception {
        Element e;
        
        // LicenseType
        transaction.appendChild(mLicenseType(fields[COL.LicenseType.ordinal()]));

        // Description
        transaction.appendChild(mDescription(fields[COL.Description.ordinal()]));

        // Territory
        transaction.appendChild(mTerritory(fields[COL.Territory.ordinal()]));

        // Start or StartCondition
        transaction.appendChild(mStart(fields[COL.Start.ordinal()]));

        // End or EndCondition
        transaction.appendChild(mEnd(fields[COL.End.ordinal()]));

        // StoreLanguage
        if ((e = mStoreLanguage(fields[COL.StoreLanguage.ordinal()])) != null)
            transaction.appendChild(e);

        // LicenseRightsDescription
        if ((e = mLicenseRightsDescription(fields[COL.LicenseRightsDescription.ordinal()])) != null)
            transaction.appendChild(e);

        // FormatProfile
        transaction.appendChild(mFormatProfile(fields[COL.FormatProfile.ordinal()]));

        // ContractID
        if ((e = mGenericElement(COL.ContractID.toString(),
                                 fields[COL.ContractID.ordinal()], false)) != null)
            transaction.appendChild(e);

        // ------------------ Term(s)
        // PriceType term
        transaction.appendChild(mPriceType(fields[COL.PriceType.ordinal()],
                                           fields[COL.PriceValue.ordinal()]));

        // SuppressionLiftDate term
        // XXX validate; required for pre-orders
        String val = fields[COL.SuppressionLiftDate.ordinal()];
        if (!val.equals("")) {
            transaction.appendChild(makeEventTerm(COL.SuppressionLiftDate.toString(), normalizeDate(val)));
        }

        // Any Term
        val = fields[COL.Any.ordinal()];
        if (!val.equals("")) {
            transaction.appendChild(makeTextTerm(COL.Any.toString(), val));
        }

        // RentalDuration Term
        val = fields[COL.RentalDuration.ordinal()];
        if (!val.equals("")) {
            if ((e = makeDurationTerm(COL.RentalDuration.toString(), val)) != null)
                transaction.appendChild(e);
        }

        // WatchDuration Term
        val = fields[COL.WatchDuration.ordinal()];
        if (!val.equals("")) {
            if ((e = makeDurationTerm(COL.WatchDuration.toString(), val)) != null)
            if (e != null)
                transaction.appendChild(e);
        }

        // HoldbackLanguage Term
        val = fields[COL.HoldbackLanguage.ordinal()];
        if (!val.equals("")) {
            transaction.appendChild(makeLanguageTerm(COL.HoldbackLanguage.toString(), val));
        }

        // HoldbackExclusionLanguage Term
        val = fields[COL.HoldbackExclusionLanguage.ordinal()];
        if (!val.equals("")) {
            transaction.appendChild(makeLanguageTerm(COL.HoldbackExclusionLanguage.toString(), val));
        }

        // OtherInstructions
        if ((e = mGenericElement(COL.OtherInstructions.toString(),
                                 fields[COL.OtherInstructions.ordinal()], false)) != null)
            transaction.appendChild(e);

        return transaction;
    } /* mTransaction() */

    public Element makeAvail(Document dom) throws Exception {
        this.dom = dom;
        Element avail = dom.createElement("Avail");
        Element e;

        // ALID
        avail.appendChild(mALID(fields[COL.AvailID.ordinal()]));
        // Disposition
        avail.appendChild(mDisposition(fields[COL.EntryType.ordinal()]));
        // Licensor
        avail.appendChild(mPublisher("Licensor", fields[COL.DisplayName.ordinal()], true));
        // Service Provider
        if ((e = mPublisher("ServiceProvider", fields[COL.ServiceProvider.ordinal()], false)) != null)
            avail.appendChild(e);
        // AvailType ('single' for a Movie)
        avail.appendChild(mGenericElement("AvailType", "single", true));
        // ShortDescription
        // XXX Doc says optional, schema says mandatory
        if ((e = mGenericElement("ShortDescription", shortDesc, true)) != null)
            avail.appendChild(e);
        // Asset
        if ((e = mAssetHeader()) != null)
            avail.appendChild(e);
        // Transaction
        if ((e = mTransactionHeader()) != null)
            avail.appendChild(e);
        // Exception Flag
        if ((e = mExceptionFlag(fields[COL.ExceptionFlag.ordinal()])) != null)
            avail.appendChild(e);

        return avail;
    }

    public Movie(AvailsSheet parent, String workType, int lineNo, String[] fields) {
        super(parent, workType, lineNo, fields);
    }
}
