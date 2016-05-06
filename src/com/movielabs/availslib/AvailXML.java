package com.movielabs.availslib;

//import java.io.IOException;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.math.BigInteger;
import java.util.List;

import org.apache.logging.log4j.*;

import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import javax.xml.bind.Unmarshaller;
//import javax.xml.parsers.DocumentBuilder;
//import javax.xml.parsers.DocumentBuilderFactory;
//import javax.xml.parsers.ParserConfigurationException;



import javax.xml.datatype.Duration;
import javax.xml.datatype.XMLGregorianCalendar;

import org.w3c.dom.*;

import com.movielabs.schema.avails.v2_0.avails.AvailAssetType;
import com.movielabs.schema.avails.v2_0.avails.AvailMetadataType;
import com.movielabs.schema.avails.v2_0.avails.AvailTransType;
import com.movielabs.schema.avails.v2_0.avails.AvailType;
import com.movielabs.schema.avails.v2_0.avails.AvailListType;
import com.movielabs.schema.avails.v2_0.avails.AvailUnitMetadataType;
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

    // DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
    // DocumentBuilder builder = factory.newDocumentBuilder();
    // Document doc = builder.parse(xmlFile);
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

        initMovieSheet();
        initEpisodeSheet();

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

    protected void initMovieSheet() {
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
    }

    protected void initEpisodeSheet() {
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
    }

    protected void addMovieRow(AvailType a) {
        Row row = movieSheet.createRow(currentMovieRow++);
        int j = 0;
        Cell cell;

        AvailAssetType asset = a.getAsset().get(0);

        List<AvailTransType> transactions = a.getTransaction();
        if (transactions.size() != 1)
            logger.warn("more than one transaction");
        AvailTransType trans = transactions.get(0);

        // Display Name
        cell = row.createCell(j++);
        cell.setCellValue(a.getLicensor().getDisplayName());

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
        cell.setCellValue(a.getDisposition().getEntryType());

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

        List<AvailTransType.Term> terms = trans.getTerm();

        // PriceType & PriceValue
        // TODO need sanity check
        String SRP = "";
        String suppressionLiftDate = "";
        Duration rentalDuration = null;
        Duration watchDuration = null;
        String anyTerm = "";
        String holdbackLanguage = "";
        String holdbackExclusionLanguage = "";
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
                    suppressionLiftDate = t.getEvent();
                } else {
                    logger.warn("duplicate 'AnnounceDate' term");
                }
                break;
            case "rentalduration":
                if (count[5]++ == 0) {
                    rentalDuration = t.getDuration();
                } else {
                    logger.warn("duplicate 'RentalDuration' term");
                }
                break;
            case "watchduration":
                if (count[6]++ == 0) {
                    watchDuration = t.getDuration();
                } else {
                    logger.warn("duplicate 'WatchDuration' term");
                }
                break;
            case "any":
                if (count[7]++ == 0) {
                    anyTerm = t.getText();
                } else {
                    logger.warn("duplicate 'Any' term");
                }
                break;
            case "holdbacklanguage":
                if (count[8]++ == 0) {
                    holdbackLanguage = t.getLanguage();
                } else {
                    logger.warn("duplicate 'HoldbackLanguage' term");
                }
                break;
            case "holdbackexclusionlanguage":
                if (count[9]++ == 0) {
                    holdbackExclusionLanguage = t.getLanguage();
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
        if (a.getALID().equals("---missing---"))
            cell.setCellValue("");
        else
            cell.setCellValue(a.getALID());

        // Metadata
        cell = row.createCell(j++);

        // AltID
        cell = row.createCell(j++);
        List<AvailMetadataType.AltIdentifier> altid = metadata.getAltIdentifier();
        if (altid.size() != 1)
            logger.warn("more than one alternate identifier");
        cell.setCellValue(altid.get(0).getIdentifier());

        // SuppressionLiftDate
        cell = row.createCell(j++);
        cell.setCellValue(suppressionLiftDate);

        // SpecialPreOrderFulfillDate
        // TODO not right
        cell = row.createCell(j++);

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
        if (a.isExceptionFlag() != null && a.isExceptionFlag().booleanValue())
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
        if (rentalDuration != null)
            cell.setCellValue(rentalDuration.getHours());
   
        // WatchDuration
        cell = row.createCell(j++);
        if (watchDuration != null)
            cell.setCellValue(watchDuration.getHours());
   
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
        cell.setCellValue(anyTerm);

        // ContractID
        cell = row.createCell(j++);
        cell.setCellValue(trans.getContractID());

        // ServiceProvider
        cell = row.createCell(j++);
        if (a.getServiceProvider() != null)
            cell.setCellValue(a.getServiceProvider().getDisplayName());

        // TotalRunTime
        cell = row.createCell(j++);
        Duration runLength = metadata.getRunLength();
        if (runLength.getHours() != 0 || runLength.getMinutes() != 0 || runLength.getSeconds() != 0) {
            cell.setCellValue(String.format("%d:%02d:%02d", runLength.getHours(), runLength.getMinutes(),
                                            runLength.getSeconds()));
        }

        // HoldbackLanguage
        cell = row.createCell(j++);
        cell.setCellValue(holdbackLanguage);

        // HoldbackExclusionLanguage
        cell = row.createCell(j++);
        cell.setCellValue(holdbackExclusionLanguage);
    }

    protected void addEpisodeRow(AvailType a) {
        Row row = episodeSheet.createRow(currentEpisodeRow++);
    }

    protected void addSeasonRow(AvailType a) {
    }

    protected void handleAvail(AvailType a) {
        //System.out.println("alid=" + a.getALID() + " type= " + a.getAvailType());
        List<AvailAssetType> assets = a.getAsset();
        if (assets.size() != 1) {
            logger.warn("more than one asset");
        }
        String workType = assets.get(0).getWorkType();
        switch(workType) {
        case "Movie":
            addMovieRow(a);
            break;
        case "Episode":
            addEpisodeRow(a);
            break;
        case "Season":
            addSeasonRow(a);
            break;
        default:
            logger.warn("invalid workType: " + workType);
        }
    }
}
