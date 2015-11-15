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

public class Episode extends SheetRow {

    private enum COL {
        DisplayName ("DisplayName"),                                   //  0
        StoreLanguage ("StoreLanguage"),                               //  1
        Territory ("md:country"),                                      //  2
        WorkType ("WorkType"),                                         //  3
        EntryType ("EntryType"),                                       //  4
        SeriesTitleInternalAlias ("SeriesTitleInternalAlias"),         //  5
        SeriesTitleDisplayUnlimited ("SeriesTitleDisplayUnlimited"),   //  6
        SeasonNumber("SeasonNumber"),                                  //  7
        EpisodeNumber("EpisodeNumber"),                                //  8
        LocalizationType ("LocalizationOffering"),                     //  9
        EpisodeTitleInternalAlias ("EpisodeTitleInternalAlias"),       // 10
        EpisodeTitleDisplayUnlimited ("EpisodeTitleDisplayUnlimited"), // 11
        SeasonTitleInternalAlias ("SeasonTitleInternalAlias"),         // 12
        SeasonTitleDisplayUnlimited ("SeasonTitleDisplayUnlimited"),   // 13
        EpisodeCount("EpisodeCount"),                                  // 14
        SeasonCount("SeasonCount"),                                    // 15
        SeriesAltID ("md:Identifier"),                                 // 16
        SeasonAltID ("md:Identifier"),                                 // 17
        EpisodeAltID ("md:Identifier"),                                // 18
        CompanyDisplayCredit("CompanyDisplayCredit"),                  // 19
        LicenseType ("LicenseType"),                                   // 20
        LicenseRightsDescription ("LicenseRightsDescription"),         // 21
        FormatProfile ("FormatProfile"),                               // 22
        Start ("StartCondition"),                                      // 23
        End ("EndCondition"),                                          // 24
        SpecialPreOrderFulfillDate("SpecialPreOrderFulfillDate"),      // 25
        PriceType ("PriceType"),                                       // 26
        PriceValue ("PriceValue"),                                     // 27
        SRP ("SRP"),                                                   // 28
        Description ("Description"),                                   // 29
        OtherTerms ("OtherTerms"),                                     // 30
        OtherInstructions ("OtherInstructions"),                       // 31
        SeriesContentID ("contentID"),                                 // 32
        SeasonContentID ("contentID"),                                 // 33
        EpisodeContentID ("contentID"),                                // 34
        EpisodeProductID ("EditEIDR-S"),                               // 35
        EncodeID ("EncodeID"),                                         // 36
        AvailID ("AvailID"),                                           // 37
        Metadata ("Metadata"),                                         // 38
        SuppressionLiftDate ("AnnounceDate"),                          // 39
        ReleaseYear ("ReleaseDate"),                                   // 40
        ReleaseHistoryOriginal ("md:Date"),                            // 41
        ReleaseHistoryPhysicalHV ("md:Date"),                          // 42
        ExceptionFlag ("ExceptionFlag"),                               // 43
        RatingSystem ("md:System"),                                    // 44
        RatingValue ("md:Value"),                                      // 45
        RatingReason ("md:RatingReason"),                              // 46
        RentalDuration ("RentalDuration"),                             // 47
        WatchDuration ("WatchDuration"),                               // 48
        FixedEndDate ("FixedEndDate"),                                 // 49
        CaptionIncluded ("USACaptionsExemptionReason"),                // 50
        CaptionExemption ("USACaptionsExemptionReason"),               // 51
        Any ("Any"),                                                   // 52
        ContractID ("ContractID"),                                     // 53
        ServiceProvider ("ServiceProvider"),                           // 54
        TotalRunTime ("RunLength"),                                    // 55
        HoldbackLanguage ("HoldbackLanguage"),                         // 56
        HoldbackExclusionLanguage ("HoldbackExclusionLanguage");       // 57


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
    
    protected Element mAssetBody(Element asset) throws Exception {
        return asset;
    }

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

        // FixedEndDate Term
        val = fields[COL.FixedEndDate.ordinal()];
        if (!val.equals("")) {
            if ((e = makeEventTerm(COL.FixedEndDate.toString(), normalizeDate(val))) != null)
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
        avail.appendChild(mLicensor(fields[COL.DisplayName.ordinal()]));
        // Service Provider
        if ((e = mServiceProvider(fields[COL.ServiceProvider.ordinal()])) != null)
            avail.appendChild(e);
        // AvailType ('episode' for an Episode)
        avail.appendChild(mGenericElement("AvailType", "single", true));
        // ShortDescription
        // XXX Doc says optional, schema says mandatory
        if ((e = mGenericElement("ShortDescription", shortDesc, true)) != null)
            avail.appendChild(e);
        // Asset
        // n = mAsset(row);
        // if (n != null)
        //     avail.appendChild(n);
        // Transaction
        // n = mTransaction(row);
        // if (n != null)
        //     avail.appendChild(n);
        // Exception Flag
        if ((e = mExceptionFlag(fields[COL.ExceptionFlag.ordinal()])) != null)
            avail.appendChild(e);

        return avail;
    }

    public Episode(AvailsSheet parent, String workType, int lineNo, String[] fields) {
        super(parent, workType, lineNo, fields);
    }
}
