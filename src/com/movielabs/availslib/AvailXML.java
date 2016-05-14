/*
 * Copyright (c) 2016 MovieLabs
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
package com.movielabs.availslib;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.math.BigInteger;
import java.util.List;

import org.apache.logging.log4j.*;

//import javax.xml.parsers.DocumentBuilder;
//import javax.xml.parsers.DocumentBuilderFactory;
//import javax.xml.parsers.ParserConfigurationException;
import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import javax.xml.bind.Unmarshaller;
import javax.xml.datatype.Duration;
import javax.xml.datatype.XMLGregorianCalendar;

import org.w3c.dom.*;

import com.movielabs.schema.avails.v2_0.avails.AvailAssetType;
import com.movielabs.schema.avails.v2_0.avails.AvailEpisodeMetadataType;
import com.movielabs.schema.avails.v2_0.avails.AvailMetadataType;
import com.movielabs.schema.avails.v2_0.avails.AvailSeasonMetadataType;
import com.movielabs.schema.avails.v2_0.avails.AvailSeriesMetadataType;
import com.movielabs.schema.avails.v2_0.avails.AvailTransType;
import com.movielabs.schema.avails.v2_0.avails.AvailType;
import com.movielabs.schema.avails.v2_0.avails.AvailListType;
import com.movielabs.schema.avails.v2_0.avails.AvailUnitMetadataType;
import com.movielabs.schema.md.v2_3.md.CompanyCreditsType;
import com.movielabs.schema.md.v2_3.md.ContentRatingDetailType;
import com.movielabs.schema.md.v2_3.md.ContentRatingType;
import com.movielabs.schema.md.v2_3.md.RegionType;
import com.movielabs.schema.md.v2_3.md.ReleaseHistoryType;
import com.movielabs.schema.mdmec.v2.PublisherType;

import org.apache.poi.POIXMLProperties;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class AvailXML {

    protected String xmlFile;
    protected Logger logger;
    protected JAXBContext jaxbContext;
    protected Unmarshaller jaxbUnmarshaller = null;
    protected List<AvailType> availList;
    protected XSSFWorkbook workbook;
    protected Sheet movieSheet, episodeSheet;
    protected int currentMovieRow, currentEpisodeRow;
    
    final static String[][] movieRows = { 
        {
            "Avail", "AvailTrans", "AvailTrans", "AvailAsset", "Disposition",                  //  A- E
            "AvailMetadata", "AvailMetadata", "AvailMetadata", "AvailTrans", "AvailTrans",     //  F- J
            "AvailTrans", "AvailTrans", "AvailTrans", "AvailTrans", "AvailTrans",              //  K- O
            "AvailTrans", "AvailTrans", "AvailTrans", "AvailTrans", "AvailAsset",              //  P- T
            "AvailAsset", "AvailAsset", "Avail", "AvailAsset", "AvailMetadata",                //  U- Y
            "AvailTrans", "AvailTrans", "AvailMetadata", "AvailMetadata", "AvailMetadata",     //  Z-AD
            "Avail", "AvailMetadata", "AvailMetadata", "AvailMetadata", "AvailTrans",          // AE-AI
            "AvailTrans", "AvailMetadata", "AvailMetadata", "AvailMetadata", "AvailTrans",     // AJ-AN
            "Avail", "AvailMetadata", "AvailTrans", "AvailTrans"                               // AO-AR
        },
        {
            "DisplayName", "StoreLanguage", "Territory", "WorkType", "EntryType",              // A- E
            "TitleInternalAlias", "TitleDisplayUnlimited", "LocalizationType", "LicenseType", "LicenseRightsDescription", // F- J
            "FormatProfile", "Start", "End", "PriceType", "PriceValue",                        // K- O
            "SRP", "Description", "OtherTerms", "OtherInstructions", "ContentID",              // P- T
            "ProductID", "EncodeID", "AvailID", "Metadata", "AltID",                           // U- Y
            "SuppressionLiftDate", "SpecialPreOrderFulfillDate", "ReleaseYear", "ReleaseHistoryOriginal", "ReleaseHistoryPhysicalHV", //  Z-AD
            "ExceptionFlag", "RatingSystem", "RatingValue", "RatingReason", "RentalDuration",  // AE-AI 
            "WatchDuration", "CaptionIncluded", "CaptionExemption", "Any", "ContractID",       // AJ-AN
            "ServiceProvider", "TotalRunTime", "HoldbackLanguage", "HoldbackExclusionLanguage" // AO-AR
        }
    };

    final static String[][] episodeRows = { 
        {
            "Avail", "AvailTrans", "AvailTrans", "AvailAsset", "Disposition",                                               //  A- E
            "AvailMetadata", "AvailMetadata", "AvailMetadata", "AvailMetadata", "AvailMetadata",                            //  F- J
            "AvailMetadata", "AvailMetadata", "AvailMetadata", "AvailMetadata", "AvailMetadata",                            //  K- O
            "AvailMetadata", "AvailMetadata", "AvailMetadata", "AvailMetadata", "AvailMetadata",                            //  P- T
            "AvailTrans", "AvailTrans", "AvailTrans", "AvailTrans", "AvailTrans",                                           //  U- Y
            "AvailTrans", "AvailTrans", "AvailTrans", "AvailTrans", "AvailTrans",                                           //  Z-AD
            "AvailTrans", "AvailTrans", "AvailAsset", "AvailAsset", "AvailAsset",                                           // AE-AI
            "AvailAsset", "AvailAsset", "Avail", "AvailAsset", "AvailMetadata",                                             // AJ-AN
            "AvailMetadata", "AvailMetadata", "AvailMetadata", "Avail", "AvailMetadata",                                    // AO-AS
            "AvailMetadata", "AvailMetadata", "AvailTrans", "AvailTrans", "AvailTrans",                                     // AT-AX
            "AvailMetadata", "AvailMetadata", "AvailMetadata", "AvailTerms", "Avail",                                       // AY-BC
            "AvailMetadata", "AvailTrans", "AvailTrans"                                                                     // BD-BF
        },
        {
            "DisplayName", "StoreLanguage", "Territory", "WorkType", "EntryType",                                           //  A- E
            "SeriesTitleInternalAlias", "SeriesTitleDisplayUnlimited", "SeasonNumber", "EpisodeNumber", "LocalizationType", //  F- J
            "EpisodeTitleInternalAlias", "EpisodeTitleDisplayUnlimited", "SeasonTitleInternalAlias", "SeasonTitleDisplayUnlimited", "EpisodeCount", //  K- O
            "SeasonCount", "SeriesAltID", "SeasonAltID", "EpisodeAltID", "CompanyDisplayCredit",                            //  P- T
            "LicenseType", "LicenseRightsDescription", "FormatProfile", "Start", "End",                                     //  U- Y
            "SpecialPreOrderFulfillDate", "PriceType", "PriceValue", "SRP", "Description",                                  //  Z-AD
            "OtherTerms", "OtherInstructions", "SeriesContentID", "SeasonContentID", "EpisodeContentID",                    // AE-AI
            "EpisodeProductID", "EncodeID", "AvailID", "Metadata", "SuppressionLiftDate",                                   // AJ-AN
            "ReleaseYear", "ReleaseHistoryOriginal", "ReleaseHistoryPhysicalHV", "ExceptionFlag", "RatingSystem",           // AO-AS
            "RatingValue", "RatingReason", "RentalDuration", "WatchDuration", "FixedEndDate",                               // AT-AX
            "CaptionIncluded", "CaptionExemption", "Any", "ContractID", "ServiceProvider",                                  // AY-BC
            "TotalRunTime", "HoldbackLanguage", "HoldbackExclusionLanguage"                                                 // BD-BF
        }
    };

    public AvailXML(String xmlFile, Logger logger) throws JAXBException {
        this.logger = logger;
        JAXBElement<AvailListType> e = null;
        this.xmlFile = xmlFile;
        File file = new File(xmlFile);
        try {
            jaxbContext = JAXBContext.newInstance(AvailListType.class);
            jaxbUnmarshaller = jaxbContext.createUnmarshaller();
            e = (JAXBElement<AvailListType>) jaxbUnmarshaller.unmarshal(file);
        } catch (JAXBException ex) {
            ex.printStackTrace();
        }
        availList = e.getValue().getAvail();
    }

    public void makeSS(String ssFile) {
        workbook = new XSSFWorkbook();

        POIXMLProperties xmlProps = workbook.getProperties();    
        POIXMLProperties.CoreProperties coreProps =  xmlProps.getCoreProperties();
        coreProps.setDescription("created by availslib from '" + xmlFile + "'");

        for (AvailType a : availList)
            handleAvail(a);

        try {
            FileOutputStream out = 
                new FileOutputStream(new File(ssFile), false);
            workbook.write(out);
            out.close();
            System.out.println("Excel written successfully..");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    protected boolean movieSheetInitialized = false;
    protected void initMovieSheet() {
        if (movieSheetInitialized)
            return;
        movieSheet = workbook.createSheet("Movie");
        currentMovieRow = 0;
        for (String[] sa : movieRows) {
            int j = 0;
            Row row = movieSheet.createRow(currentMovieRow++);
            for (String s : sa) {
                Cell cell = row.createCell(j++);
                cell.setCellValue(s);
            }
        }
        movieSheetInitialized = true;
    }

    protected boolean episodeSheetInitialized = false;
    protected void initEpisodeSheet() {
        if (episodeSheetInitialized)
            return;
        episodeSheet = workbook.createSheet("Episode");
        currentEpisodeRow = 0;
        for (String[] sa : episodeRows) {
            int j = 0;
            Row row = episodeSheet.createRow(currentEpisodeRow++);
            for (String s : sa) {
                Cell cell = row.createCell(j++);
                cell.setCellValue(s);
            }
        }
        episodeSheetInitialized = true;
    }

    protected int createAtoE(int j, Row row, AvailType avail, AvailAssetType asset, AvailTransType trans) {
        Cell cell;
        
        // Display Name
        cell = row.createCell(j++);
        cell.setCellValue(avail.getLicensor().getDisplayName());

        // Store Language
        List<String> storeLang = trans.getStoreLanguage();
        if (storeLang.size() != 1)
            logger.warn("more than one store language");
        cell = row.createCell(j++);
        cell.setCellValue(storeLang.get(0));

        // Territory
        List<RegionType> territory = trans.getTerritory();
        if (territory.size() != 1)
            logger.warn("more than one territory");
        cell = row.createCell(j++);
        cell.setCellValue(territory.get(0).getCountry());

        // WorkType
        cell = row.createCell(j++);
        cell.setCellValue(asset.getWorkType());

        // EntryType
        cell = row.createCell(j++);
        cell.setCellValue(avail.getDisposition().getEntryType());

        return j;
    } /* createAtoE() */

    protected int createFtoH(int j, Row row, AvailSeriesMetadataType seriesMetadata, AvailSeasonMetadataType seasonMetadata) {
        Cell cell;

        // SeriesTitleInternalAlias (F)
        cell = row.createCell(j++);
        cell.setCellValue(seriesMetadata.getSeriesTitleInternalAlias());

        // SeriesTitleDisplayUnlimited (G)
        cell = row.createCell(j++);
        cell.setCellValue(seriesMetadata.getSeriesTitleDisplayUnlimited());

        // SeasonNumber (H)
        cell = row.createCell(j++);
        cell.setCellValue(seasonMetadata.getSeasonNumber().getNumber());

        return j;
    }

    protected int createMtoR(int j, Row row, AvailSeriesMetadataType seriesMetadata, AvailSeasonMetadataType seasonMetadata) {
        Cell cell;

        // SeasonTitleInternalAlias (M)
        cell = row.createCell(j++);
        cell.setCellValue(seasonMetadata.getSeasonTitleInternalAlias());

        // SeasonTitleDisplayUnlimited (N)
        cell = row.createCell(j++);
        cell.setCellValue(seasonMetadata.getSeasonTitleDisplayUnlimited());

        // EpisodeCount (O)
        AvailSeasonMetadataType.NumberOfEpisodes nOE = seasonMetadata.getNumberOfEpisodes();
        if (nOE != null) {
            cell = row.createCell(j++);
            cell.setCellValue(nOE.getValue().intValue());
        } else {
            j++;
        }

        // SeasonCount (P)
        AvailSeriesMetadataType.NumberOfSeasons nOS = seriesMetadata.getNumberOfSeasons();
        if (nOS != null) {
            cell = row.createCell(j++);
            cell.setCellValue(nOS.getValue().intValue());
        } else {
            j++;
        }

        // SeriesAltID (Q)
        List<AvailSeriesMetadataType.SeriesAltIdentifier> seriesAltId = seriesMetadata.getSeriesAltIdentifier();
        if (seriesAltId != null && seriesAltId.size() == 1) {
            cell = row.createCell(j++);
            cell.setCellValue(seriesAltId.get(0).getIdentifier());
        } else {
            j++;
        }

        // SeasonAltID (R)
        List<AvailSeasonMetadataType.SeasonAltIdentifier> seasonAltId = seasonMetadata.getSeasonAltIdentifier();
        if (seasonAltId != null && seasonAltId.size() == 1) {
            cell = row.createCell(j++);
            cell.setCellValue(seasonAltId.get(0).getIdentifier());
        } else {
            j++;
        }

        return j;
    } /* createMtoR() */

    protected int createT(int j, Row row, AvailSeriesMetadataType seriesMetadata) {
        Cell cell;

        // CompanyDisplayCredit (T)
        List<CompanyCreditsType> credits = seriesMetadata.getCompanyDisplayCredit();
        if (credits.size() == 1) {
            List<CompanyCreditsType.DisplayString> displayStrings = credits.get(0).getDisplayString();
            if (credits.size() == 1) {
                cell = row.createCell(j++);
                cell.setCellValue(displayStrings.get(0).getValue());
            }
        } else { // TODO warnings
            j++;
        }
 
        return j;
    }

    protected int createTransactions1(int j, Row row, AvailTransType trans) {
        Cell cell;

        // LicenseType
        cell = row.createCell(j++);
        cell.setCellValue(trans.getLicenseType());
        
        // LicenseRightsDescription
        cell = row.createCell(j++);
        cell.setCellValue(trans.getLicenseRightsDescription());
        
        // FormatProfile
        cell = row.createCell(j++);
        cell.setCellValue(trans.getFormatProfile());

        // Start
        cell = row.createCell(j++);
        XMLGregorianCalendar gc = trans.getStart();
        cell.setCellValue(String.format("%04d-%02d-%02d", gc.getYear(), gc.getMonth(), gc.getDay()));

        // End
        cell = row.createCell(j++);
        gc = trans.getEnd();
        if (gc != null) {
            cell.setCellValue(String.format("%04d-%02d-%02d", gc.getYear(), gc.getMonth(), gc.getDay()));
        } else {
            cell.setCellValue(trans.getEndCondition());
        }

        return j;
    } /* createTransactions1() */

    private class Hack {
        public String suppressionLiftDate;
        public Duration rentalDuration;
        public Duration watchDuration;
        public String anyTerm;
        public String holdbackLanguage;
        public String holdbackExclusionLanguage;
    }

    protected int createTransactions2(int j, Row row, AvailTransType trans, Hack h) {
        Cell cell;

        List<AvailTransType.Term> terms = trans.getTerm();

        // PriceType & PriceValue
        // TODO need sanity check
        String SRP = "";
        h.suppressionLiftDate = "";
        h.rentalDuration = null;
        h.watchDuration = null;
        h.anyTerm = "";
        h.holdbackLanguage = "";
        h.holdbackExclusionLanguage = "";
        int count[] = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0};
        for (AvailTransType.Term t : terms) { // this is n^2 but there aren't many terms so big deal
            String name = t.getTermName();
            switch(name.toLowerCase()) {
            case "tier":
                if (count[0]++ == 0) {
                    cell = row.createCell(j++);
                    cell.setCellValue(name);
                    cell = row.createCell(j++);
                    double d = Double.parseDouble(t.getText());
                    cell.setCellValue(String.format("%.0f", d));
                } else {
                    logger.warn("duplicate 'tier' term");
                }
                break;
            case "category":
                if (count[1]++ == 0) {
                    cell = row.createCell(j++);
                    cell.setCellValue(name);
                    cell = row.createCell(j++);
                    cell.setCellValue(t.getText());
                } else {
                    logger.warn("duplicate 'category' term");
                }
                break;
            case "wsp":
                if (count[2]++ == 0) {
                    cell = row.createCell(j++);
                    cell.setCellValue(name);
                    cell = row.createCell(j++);
                    cell.setCellValue(String.format("%.2f", t.getMoney().getValue().doubleValue()));
                } else {
                    logger.warn("duplicate 'WSP' term");
                }
                break;
            case "srp":
                if (count[3]++ == 0) {
                    SRP = String.format("%.2f", t.getMoney().getValue().doubleValue());  // Extract and Save
                } else {
                    logger.warn("duplicate 'SRP' term");
                }
                break;
            case "announcedate":
                if (count[4]++ == 0) {
                    h.suppressionLiftDate = t.getEvent();
                } else {
                    logger.warn("duplicate 'AnnounceDate' term");
                }
                break;
            case "rentalduration":
                if (count[5]++ == 0) {
                    h.rentalDuration = t.getDuration();
                } else {
                    logger.warn("duplicate 'RentalDuration' term");
                }
                break;
            case "watchduration":
                if (count[6]++ == 0) {
                    h.watchDuration = t.getDuration();
                } else {
                    logger.warn("duplicate 'WatchDuration' term");
                }
                break;
            case "any":
                if (count[7]++ == 0) {
                    h.anyTerm = t.getText();
                } else {
                    logger.warn("duplicate 'Any' term");
                }
                break;
            case "holdbacklanguage":
                if (count[8]++ == 0) {
                    h.holdbackLanguage = t.getLanguage();
                } else {
                    logger.warn("duplicate 'HoldbackLanguage' term");
                }
                break;
            case "holdbackexclusionlanguage":
                if (count[9]++ == 0) {
                    h.holdbackExclusionLanguage = t.getLanguage();
                } else {
                    logger.warn("duplicate 'HoldbackExclusionLanguage' term");
                }
                break;
            default:
                logger.warn("unknown term: ", t.getTermName());
            }
        }

        // SRP
        cell = row.createCell(j++);
        cell.setCellValue(SRP);

        // Description
        cell = row.createCell(j++);
        if (trans.getDescription().equals("---missing---"))
            cell.setCellValue("");
        else
            cell.setCellValue(trans.getDescription());

        // Other Terms
        // TODO not right
        cell = row.createCell(j++);
        cell.setCellValue("");

        // OtherInstructions
        cell = row.createCell(j++);
        cell.setCellValue(trans.getOtherInstructions());

        return j;
    } /* createTransactions2() */

    protected void addMovieRow(AvailType avail) {
        Row row = movieSheet.createRow(currentMovieRow++);
        int j = 0;
        Cell cell;

        AvailAssetType asset = avail.getAsset().get(0);

        List<AvailTransType> transactions = avail.getTransaction();
        if (transactions.size() != 1)
            logger.warn("more than one transaction");
        AvailTransType trans = transactions.get(0);

        j = createAtoE(j, row, avail, asset, trans);

        // TitleInternalAlias
        AvailMetadataType metadata = asset.getMetadata();
        cell = row.createCell(j++);
        cell.setCellValue(metadata.getTitleInternalAlias());

        // TitleDisplayUnlimited
        cell = row.createCell(j++);
        cell.setCellValue(metadata.getTitleDisplayUnlimited().getValue());

        // LocalizationType
        cell = row.createCell(j++);
        cell.setCellValue(metadata.getLocalizationOffering());

        // I-M common
        j = createTransactions1(j, row, trans);

        Hack h = new Hack();

        // N-S common
        j = createTransactions2(j, row, trans, h);

        // ContentID
        cell = row.createCell(j++);
        cell.setCellValue(asset.getContentID());

        // ProductID
        cell = row.createCell(j++);
        cell.setCellValue(metadata.getEditEIDRS());

        // EncodeID
        cell = row.createCell(j++);
        cell.setCellValue(metadata.getEncodeID());
 
        // AvailID
        cell = row.createCell(j++);
        if (avail.getALID().equals("---missing---"))
            cell.setCellValue("");
        else
            cell.setCellValue(avail.getALID());

        // Metadata
        cell = row.createCell(j++);

        // AltID
        cell = row.createCell(j++);
        List<AvailMetadataType.AltIdentifier> altId = metadata.getAltIdentifier();
        if (altId != null && altId.size() == 1)
            cell.setCellValue(altId.get(0).getIdentifier());
        else
            logger.warn("more than one or null alternate identifier");

        // SuppressionLiftDate
        cell = row.createCell(j++);
        cell.setCellValue(h.suppressionLiftDate);

        // SpecialPreOrderFulfillDate
        // TODO not right
        j++;

        // ReleaseYear
        cell = row.createCell(j++);
        cell.setCellValue(metadata.getReleaseDate());

        // ReleaseHistoryOriginal
        String releaseHistoryPhysicalHV = "";
        String releaseHistoryOriginal = "";
        cell = row.createCell(j++);
        List<ReleaseHistoryType> relHistory = metadata.getReleaseHistory();
        for (ReleaseHistoryType rh : relHistory) {
            switch(rh.getReleaseType().getValue()) {
            case "DVD":
                releaseHistoryPhysicalHV = rh.getDate().getValue();
                break;
            case "original":
                releaseHistoryOriginal = rh.getDate().getValue();
                break;
            }
        }
        cell.setCellValue(releaseHistoryOriginal);

        // ReleaseHistoryPhysicalHV
        cell = row.createCell(j++);
        cell.setCellValue(releaseHistoryPhysicalHV);

        // ExceptionFlag
        cell = row.createCell(j++);
        if (avail.isExceptionFlag() != null && avail.isExceptionFlag().booleanValue())
            cell.setCellValue("YES");
        else
            cell.setCellValue("");

        ContentRatingType ratings = metadata.getRatings();
        if (ratings == null) {
            j += 3;
        } else {
            List<ContentRatingDetailType> rlist = ratings.getRating();
            if (rlist.size() != 1)
                logger.warn("more than one rating specified");
            // RatingSystem
            cell = row.createCell(j++);
            cell.setCellValue(rlist.get(0).getSystem());
            
            // RatingValue
            cell = row.createCell(j++);
            cell.setCellValue(rlist.get(0).getValue());

            // RatingReason
            cell = row.createCell(j++);
            List<String> reasons = rlist.get(0).getReason();
            int i = 0;
            String s = "";
            int tmp = reasons.size();
            for (String r : reasons) {
                if (i++ > 0)
                    s += ",";
                s += r;
            }
            cell.setCellValue(s);
        }

        // RentalDuration
        cell = row.createCell(j++);
        if (h.rentalDuration != null)
            cell.setCellValue(h.rentalDuration.getHours());
   
        // WatchDuration
        cell = row.createCell(j++);
        if (h.watchDuration != null)
            cell.setCellValue(h.watchDuration.getHours());
   
        BigInteger exemption = metadata.getUSACaptionsExemptionReason();
        if (exemption != null) {
            cell = row.createCell(j++);
            cell.setCellValue("No");                 // CaptionIncluded
            cell = row.createCell(j++);
            cell.setCellValue(exemption.intValue()); // CaptionExemption
        } else {
            cell = row.createCell(j++);
            cell.setCellValue("Yes");                // CaptionIncluded
            j++;                                     // CaptionExemption
        }

        // Any
        cell = row.createCell(j++);
        cell.setCellValue(h.anyTerm);

        // ContractID
        cell = row.createCell(j++);
        cell.setCellValue(trans.getContractID());

        // ServiceProvider
        cell = row.createCell(j++);
        if (avail.getServiceProvider() != null)
            cell.setCellValue(avail.getServiceProvider().getDisplayName());

        // TotalRunTime
        cell = row.createCell(j++);
        Duration runLength = metadata.getRunLength();
        if (runLength.getHours() != 0 || runLength.getMinutes() != 0 || runLength.getSeconds() != 0) {
            cell.setCellValue(String.format("%d:%02d:%02d", runLength.getHours(), runLength.getMinutes(),
                                            runLength.getSeconds()));
        }

        // HoldbackLanguage
        cell = row.createCell(j++);
        cell.setCellValue(h.holdbackLanguage);

        // HoldbackExclusionLanguage
        cell = row.createCell(j++);
        cell.setCellValue(h.holdbackExclusionLanguage);
    } /* addMovieRow() */

    protected void addEpisodeRow(AvailType avail) {
        Row row = episodeSheet.createRow(currentEpisodeRow++);
        int j = 0;
        Cell cell;

        AvailAssetType asset = avail.getAsset().get(0);

        List<AvailTransType> transactions = avail.getTransaction();
        if (transactions.size() != 1)
            logger.warn("more than one transaction");
        AvailTransType trans = transactions.get(0);

        j = createAtoE(j, row, avail, asset, trans);

        AvailEpisodeMetadataType episodeMetadata = asset.getEpisodeMetadata();
        AvailSeasonMetadataType  seasonMetadata  = episodeMetadata.getSeasonMetadata();
        AvailSeriesMetadataType  seriesMetadata  = seasonMetadata.getSeriesMetadata();

        j = createFtoH(j, row, seriesMetadata, seasonMetadata);

        // EpisodeNumber (I)
        cell = row.createCell(j++);
        cell.setCellValue(episodeMetadata.getEpisodeNumber().getNumber());

        // LocalizationType (J)
        cell = row.createCell(j++);
        cell.setCellValue(episodeMetadata.getLocalizationOffering());

        // EpisodeTitleInternalAlias (K)
        cell = row.createCell(j++);
        cell.setCellValue(episodeMetadata.getTitleInternalAlias());

        // EpisodeTitleDisplayUnlimited (L)
        AvailMetadataType.TitleDisplayUnlimited tDU = episodeMetadata.getTitleDisplayUnlimited();
        if (tDU != null) {
            cell = row.createCell(j++);
            cell.setCellValue(tDU.getValue());
        } else {
            j++;
        }

        // M-R common
        j = createMtoR(j, row, seriesMetadata, seasonMetadata);

        // EpisodeAltID (S)
        List<AvailMetadataType.AltIdentifier> episodeAltId = episodeMetadata.getAltIdentifier();
        if (episodeAltId != null && episodeAltId.size() == 1) {
            cell = row.createCell(j++);
            cell.setCellValue(episodeAltId.get(0).getIdentifier());
        } else {
            j++;
        }

        // CompanyDisplayCredit (T)
        j = createT(j, row, seriesMetadata);

        // U-Y common
        j = createTransactions1(j, row, trans);

        // SpecialPreOrderFulfillDate (Z)
        // TODO not right
        j++;

        // AA-AF common
        Hack h = new Hack();
        j = createTransactions2(j, row, trans, h);

    } /* addEpisodeRow() */

    protected void addSeasonRow(AvailType avail) {
        Row row = episodeSheet.createRow(currentEpisodeRow++);
        int j = 0;
        Cell cell;

        AvailAssetType asset = avail.getAsset().get(0);

        List<AvailTransType> transactions = avail.getTransaction();
        if (transactions.size() != 1)
            logger.warn("more than one transaction");
        AvailTransType trans = transactions.get(0);

        j = createAtoE(j, row, avail, asset, trans);

        AvailSeasonMetadataType seasonMetadata = asset.getSeasonMetadata();
        AvailSeriesMetadataType seriesMetadata = seasonMetadata.getSeriesMetadata();

        j = createFtoH(j, row, seriesMetadata, seasonMetadata);

        // no EpisodeNumber, LocalizationType, EpisodeTitleInternalAlias, or EpisodeTitleDisplayUnlimited (I-L)
        j += 4;

        // M-R common
        j = createMtoR(j, row, seriesMetadata, seasonMetadata);

        // no EpisodeAltID (S)
        j++;

        // CompanyDisplayCredit (T)
        j = createT(j, row, seriesMetadata);

        // U-Y common
        j = createTransactions1(j, row, trans);

        // SpecialPreOrderFulfillDate (Z)
        // TODO not right
        j++;

        // AA-AF common
        Hack h = new Hack();
        j = createTransactions2(j, row, trans, h);

    } /* addSeasonRow() */

    protected void handleAvail(AvailType avail) {
        //System.out.println("alid=" + a.getALID() + " type= " + a.getAvailType());
        List<AvailAssetType> assets = avail.getAsset();
        if (assets.size() != 1) {
            logger.warn("more than one asset");
        }
        String workType = assets.get(0).getWorkType();
        switch(workType) {
        case "Movie":
            initMovieSheet();
            addMovieRow(avail);
            break;
        case "Episode":
            initEpisodeSheet();
            addEpisodeRow(avail);
            break;
        case "Season":
            initEpisodeSheet();
            addSeasonRow(avail);
            break;
        default:
            logger.warn("invalid workType: " + workType);
        }
    }
}
