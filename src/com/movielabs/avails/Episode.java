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
        EpisodeTitleInternalAlias ("TitleInternalAlias"),              // 10
        EpisodeTitleDisplayUnlimited ("TitleDisplayUnlimited"),        // 11
        SeasonTitleInternalAlias ("SeasonTitleInternalAlias"),         // 12
        SeasonTitleDisplayUnlimited ("SeasonTitleDisplayUnlimited"),   // 13
        EpisodeCount("NumberOfEpisodes"),                              // 14
        SeasonCount("NumberOfSeasons"),                                // 15
        SeriesAltID ("SeriesAltIdentifier"),                           // 16
        SeasonAltID ("SeasonAltIdentifier"),                           // 17
        EpisodeAltID ("AltIdentifier"),                                // 18
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
        SeriesContentID ("SeriesContentID"),                           // 32
        SeasonContentID ("SeasonContentID"),                           // 33
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
    
    private void TDUHelper(Element metadata, COL col, COL rcol, String sub) throws Exception {
        Element e;
        if (sub != null) {
            if (fields[col.ordinal()].equals("") && fields[rcol.ordinal()].equals(""))
                fields[col.ordinal()] = sub;
        }
        if (fields[col.ordinal()].equals(""))
            fields[col.ordinal()] = fields[rcol.ordinal()];
        if ((e = mGenericElement(col.toString(), 
                                 fields[col.ordinal()], false)) != null)
            metadata.appendChild(e);
    }

    private void altIDHelper(Element metadata, COL col) throws Exception {
        if (!fields[col.ordinal()].equals("")) { // optional
            Element altID = dom.createElement(col.toString());
            Element cid = dom.createElement("md:Namespace");
            Text tmp = dom.createTextNode(MISSING);
            cid.appendChild(tmp);
            altID.appendChild(cid);
            altID.appendChild(mGenericElement("md:Identifier", 
                                              fields[col.ordinal()], false));
            Element loc = dom.createElement("md:Location");
            tmp = dom.createTextNode(MISSING);
            loc.appendChild(tmp);
            altID.appendChild(loc);
            metadata.appendChild(altID);
        }
    }

    protected Element mAssetBody(Element asset) throws Exception {
        Element e;
        Attr attr;
        Text tmp;
        String territory = fields[COL.Territory.ordinal()];
        
        // ------------------------------- Episode

        // EpisodeContentID --> Asset@ContentID
        Element episodeMetadata = dom.createElement("EpisodeMetadata");
        attr = dom.createAttribute(COL.EpisodeContentID.toString());
        attr.setValue(fields[COL.EpisodeContentID.ordinal()]);
        asset.setAttributeNode(attr);

        // EpisodeTitleDisplayUnlimited
        // XXX Optional in SS, Required in XML; workaround by assigning it internal alias value
        TDUHelper(episodeMetadata, COL.EpisodeTitleDisplayUnlimited, COL.EpisodeTitleInternalAlias,
                  null);

        // TitleInternalAlias
        episodeMetadata.appendChild(mGenericElement(COL.EpisodeTitleInternalAlias.toString(),
                                             fields[COL.EpisodeTitleInternalAlias.ordinal()], true));

        // ProductID --> EditEIDR-S
        if (!fields[COL.EpisodeProductID.ordinal()].equals("")) { // optional field
            String productID = normalizeEIDR(fields[COL.EpisodeProductID.ordinal()]);
            if (productID == null) {
                reportError("Invalid ProductID: " + fields[COL.EpisodeProductID.ordinal()]);
            } else {
                fields[COL.EpisodeProductID.ordinal()] = productID;
            }
            episodeMetadata.appendChild(mGenericElement(COL.EpisodeProductID.toString(),
                                                 fields[COL.EpisodeProductID.ordinal()], false));
        }

        // EpisodeAltID --> AltIdentifier
        altIDHelper(episodeMetadata, COL.EpisodeAltID);

        // ReleaseYear ---> ReleaseDate
        if (!fields[COL.ReleaseYear.ordinal()].equals("")) { // optional
            String year = normalizeYear(fields[COL.ReleaseYear.ordinal()]);
            if (year == null) {
                reportError("Invalid ReleaseYear: " + fields[COL.ReleaseYear.ordinal()]);
            } else {
                fields[COL.ReleaseYear.ordinal()] = year;
            }
            episodeMetadata.appendChild(mGenericElement(COL.ReleaseYear.toString(), 
                                                 fields[COL.ReleaseYear.ordinal()], false));
        }

        // TotalRunTime --> RunLength
        // XXX TotalRunTime is optional, RunLength is required
        if ((e = mRunLength(fields[COL.TotalRunTime.ordinal()])) != null)
            episodeMetadata.appendChild(e);

        // ReleaseHistoryOriginal ---> ReleaseHistory/Date
        String date = fields[COL.ReleaseHistoryOriginal.ordinal()];
        if (!date.equals("")) {
            if ((e = mReleaseHistory(COL.ReleaseHistoryOriginal.toString(),
                                     normalizeDate(date), "Broadcast")) != null)
                episodeMetadata.appendChild(e);
        }
         
        // ReleaseHistoryPhysicalHV ---> ReleaseHistory/Date
        date = fields[COL.ReleaseHistoryPhysicalHV.ordinal()];
        if (!date.equals("")) {
            if ((e = mReleaseHistory(COL.ReleaseHistoryPhysicalHV.toString(),
                                     normalizeDate(date), "DVD")) != null)
                episodeMetadata.appendChild(e);
        }

        // CaptionIncluded/CaptionException
        mCaption(episodeMetadata, fields[COL.CaptionIncluded.ordinal()],
                 fields[COL.CaptionExemption.ordinal()], territory);

        // RatingSystem ---> Ratings
        mRatings(episodeMetadata, fields[COL.RatingSystem.ordinal()], 
                 fields[COL.RatingValue.ordinal()], fields[COL.RatingReason.ordinal()], 
                 territory);

        // EncodeID --> EditEIDR-S
        if (!fields[COL.EncodeID.ordinal()].equals("")) { // optional field
            String encodeID = normalizeEIDR(fields[COL.EncodeID.ordinal()]);
            if (encodeID == null) {
                reportError("Invalid EncodeID: " + fields[COL.EncodeID.ordinal()]);
            } else {
                fields[COL.EncodeID.ordinal()] = encodeID;
            }
            episodeMetadata.appendChild(mGenericElement(COL.EncodeID.toString(),
                                                 fields[COL.EncodeID.ordinal()], false));
        }

        // LocalizationType --> LocalizationOffering
        if ((e = mLocalizationType(episodeMetadata, fields[COL.LocalizationType.ordinal()])) != null) {
            episodeMetadata.appendChild(e);
        }

        // EpisodeNumber
        episodeMetadata.appendChild(mCount(COL.EpisodeNumber.toString(),
                                           fields[COL.EpisodeNumber.ordinal()]));
                                                           
        // ------------------------------- Episode/Season
        Element seasonMetadata = dom.createElement("SeasonMetadata");

        if ((e = mGenericElement(COL.SeasonContentID.toString(),
                                 fields[COL.SeasonContentID.ordinal()], false)) != null)
            seasonMetadata.appendChild(e);

        // SeasonDisplayUnlimited
        // XXX Optional in SS, Required in XML; workaround by assigning it internal alias value
        TDUHelper(seasonMetadata, COL.SeasonTitleDisplayUnlimited, COL.SeasonTitleInternalAlias,
                  MISSING);

        // SeasonTitleInternalAlias
        seasonMetadata.appendChild(mGenericElement(COL.SeasonTitleInternalAlias.toString(),
                                             fields[COL.SeasonTitleInternalAlias.ordinal()], true));

        // SeasonNumber
        seasonMetadata.appendChild(mCount(COL.SeasonNumber.toString(),
                                           fields[COL.SeasonNumber.ordinal()]));

        // SeasonAltIdentifier
        altIDHelper(seasonMetadata, COL.SeasonAltID);

        // EpisodeCount --> NumberOfEpisodes
        String nEp = fields[COL.EpisodeCount.ordinal()];
        if (!nEp.equals("")) {
            int n = normalizeInt(nEp);
            e = dom.createElement(COL.EpisodeCount.toString());
            tmp = dom.createTextNode(String.valueOf(n));
            e.appendChild(tmp);
            seasonMetadata.appendChild(e);
        }

        episodeMetadata.appendChild(seasonMetadata);

        // ------------------------------- Episode/Season/Series
        Element seriesMetadata = dom.createElement("SeriesMetadata");
        e = dom.createElement(COL.SeriesContentID.toString());
        tmp = dom.createTextNode(fields[COL.SeriesContentID.ordinal()]);
        e.appendChild(tmp);
        seriesMetadata.appendChild(e);

        // SeriesDisplayUnlimited
        // XXX Optional in SS, Required in XML; workaround by assigning it internal alias value
        TDUHelper(seriesMetadata, COL.SeriesTitleDisplayUnlimited, COL.SeriesTitleInternalAlias,
                  MISSING);

        // SeriesInternalAlias
        seriesMetadata.appendChild(mGenericElement(COL.SeriesTitleInternalAlias.toString(),
                                             fields[COL.SeriesTitleInternalAlias.ordinal()], true));

        altIDHelper(seriesMetadata, COL.SeriesAltID);

        // SeasonCount --> NumberOfSeasons
        String nSe = fields[COL.SeasonCount.ordinal()];
        if (!nSe.equals("")) {
            int n = normalizeInt(nSe);
            e = dom.createElement(COL.SeasonCount.toString());
            tmp = dom.createTextNode(String.valueOf(n));
            e.appendChild(tmp);
            seriesMetadata.appendChild(e);
        }

        seasonMetadata.appendChild(seriesMetadata);

        asset.appendChild(episodeMetadata);
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

        // XXX currency not specified
        String val = fields[COL.SRP.ordinal()];
        if (!val.equals(""))
            transaction.appendChild(makeMoneyTerm("SRP", fields[COL.SRP.ordinal()], null));

        // SuppressionLiftDate term
        // XXX validate; required for pre-orders
        val = fields[COL.SuppressionLiftDate.ordinal()];
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
        avail.appendChild(mPublisher("Licensor", fields[COL.DisplayName.ordinal()], true));

        // Service Provider
        if ((e = mPublisher("ServiceProvider", fields[COL.ServiceProvider.ordinal()], false)) != null)
            avail.appendChild(e);
        // AvailType ('episode' for an Episode)
        avail.appendChild(mGenericElement("AvailType", "episode", true));

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

    public Episode(AvailsSheet parent, String workType, int lineNo, String[] fields) {
        super(parent, workType, lineNo, fields);
    }
}
