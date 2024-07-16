package com.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataCleaner {

    private static final Map<String, String> customer_names = new HashMap<>();
    
    static {
        customer_names.put("ablcn", "ABL-WUH-Ocean");
        customer_names.put("ablcn/kylin", "ABL-WUH-Ocean");
        customer_names.put("abl-cn", "ABL-WUH-Ocean");
        customer_names.put("afsky", "AF Sky");
        customer_names.put("airsea", "Air Sea Transport");
        customer_names.put("alexbest", "alexbest");
        customer_names.put("americannew", "American New");
        customer_names.put("ano", "Ano 前海安诺");
        customer_names.put("apex", "Apex");
        customer_names.put("apexlc", "Apex");
        customer_names.put("apexwh", "Apex");
        customer_names.put("apex-amz", "Apex");
        customer_names.put("apex-lc", "Apex");
        customer_names.put("aspeed", "ASPEED/RAPID KIND/SXJD");
        customer_names.put("aspeed-sxjd", "ASPEED/RAPID KIND/SXJD");
        customer_names.put("8dt", "Badatong 永利八达通");
        customer_names.put("bluestar", "Bluestar— —星蓝国际");
        customer_names.put("bluestars", "Bluestar— —星蓝国际");
        customer_names.put("cheetah", "CHEETAH");
        customer_names.put("vastlog", "China Vast Logistics Co.,Ltd.");
        customer_names.put("长河航通-创辰", "Chuangchen-创辰");
        customer_names.put("cinco", "Cinco");
        customer_names.put("cjexpress", "CJ Express");
        customer_names.put("wuh-cts", "CTS");
        customer_names.put("dimerco_guangzhou", "Dimerco");
        customer_names.put("egobike", "Ego Bike USA");
        customer_names.put("egobike-dap-lequmod", "Ego Bike USA");
        customer_names.put("egobike-lequmodinc", "Ego Bike USA");
        customer_names.put("elite", "Elite Freight");
        customer_names.put("blistex", "Eman/Blistex");
        customer_names.put("eman", "Eman/Blistex");
        customer_names.put("eman/blistex", "Eman/Blistex");
        customer_names.put("eman/grandma", "Eman/Blistex");
        customer_names.put("emotrans", "Emo Trans");
        customer_names.put("evangel", "Evangel Shipping Inc");
        customer_names.put("fcc", "FCC");
        customer_names.put("finwhale", "Fin Whale");
        customer_names.put("fmglobal", "FM Global Logistics");
        customer_names.put("forest-andy", "Forest");
        customer_names.put("forest-angel", "Forest");
        customer_names.put("forest-camila", "Forest");
        customer_names.put("forest-camille", "Forest");
        customer_names.put("forest-hades", "Forest");
        customer_names.put("forest-helen", "Forest");
        customer_names.put("forest-kaylee", "Forest");
        customer_names.put("forest-nicole", "Forest");
        customer_names.put("forest-tessa", "Forest");
        customer_names.put("forest-yvonne", "Forest");
        customer_names.put("forest-helene", "Forest");
        customer_names.put("forest-nancy", "Forest");
        customer_names.put("forest-nora", "Forest");
        customer_names.put("forest-rachel", "Forest");
        customer_names.put("forest-shirley", "Forest");
        customer_names.put("fortunetau", "Fortune Tau");
        customer_names.put("gce/transpacific", "GCE Group");
        customer_names.put("globalfreight", "Global Freight");
        customer_names.put("globalpass", "Global Pass/GTS");
        customer_names.put("globalpass-gts", "Global Pass/GTS");
        customer_names.put("goldenarcus", "Golden Arcus");
        customer_names.put("grandfreight鹏远", "Grand Freight Service");
        customer_names.put("greenlite", "Greenlite");
        customer_names.put("haiyi", "Hai Yi Global INC");
        customer_names.put("hongyu", "Hongyu 虹玉");
        customer_names.put("igt", "IGT");
        customer_names.put("jeny", "JENY");
        customer_names.put("9fang", "Jiu Fang / 九方 / Tofba");
        customer_names.put("jiufang/tofba", "Jiu Fang / 九方 / Tofba");
        customer_names.put("jusda", "JUSDA");
        customer_names.put("kaiya", "KAI YA");
        customer_names.put("ningbokaiyu", "KAI YU");
        customer_names.put("kenpax", "Kenpax");
        customer_names.put("karken", "Kraken Logistics");
        customer_names.put("kraken", "Kraken Logistics");
        customer_names.put("krakenlogistics", "Kraken Logistics");
        customer_names.put("lcgroup", "LC");
        customer_names.put("leader", "Leaders Air Sea");
        customer_names.put("leaders", "Leaders Air Sea");
        customer_names.put("leadersair&sea", "Leaders Air Sea");
        customer_names.put("linkw", "Link W");
        customer_names.put("linktrans", "LinkTrans/Shenzhen Kok");
        customer_names.put("link-trans", "LinkTrans/Shenzhen Kok");
        customer_names.put("ablcn/litelogo", "Litelogo");
        customer_names.put("litelogo", "Litelogo");
        customer_names.put("meiwo", "MEIWO/ONE DREAM");
        customer_names.put("mile", "Mile Intl/Ying Li");
        customer_names.put("mileinternational", "Mile Intl/Ying Li");
        customer_names.put("mz", "MingZhi 铭志");
        customer_names.put("neptune", "Neptune Shipping Limited");
        customer_names.put("nbjd", "NingBo Jingda");
        customer_names.put("ningbojingda", "NingBo Jingda");
        customer_names.put("niuku", "NIUKU");
        customer_names.put("oec", "OEC / 海硕");
        customer_names.put("olwarehouse", "OL Warehouse-Hugo");
        customer_names.put("penavico-spey", "Penavico Ningbo");
        customer_names.put("pglexport", "PGL");
        customer_names.put("pgnetwork", "PG-Net 传鸽");
        customer_names.put("pg-network", "PG-Net 传鸽");
        customer_names.put("pigeonnetwork", "PG-Net 传鸽");
        customer_names.put("pico", "Pico Logistics");
        customer_names.put("pico/ddp", "Pico Logistics");
        customer_names.put("plj", "PLJ Link-Ever");
        customer_names.put("pljlinkever", "PLJ Link-Ever");
        customer_names.put("plj-linkever", "PLJ Link-Ever");
        customer_names.put("qingdaozhexin", "Qingdao Zhexin");
        customer_names.put("zhexing", "Qingdao Zhexin");
        customer_names.put("rapidfreight", "Rapid Freight");
        customer_names.put("weitu", "Rayway International Logistics-weitu");
        customer_names.put("xiamenweitu", "Rayway International Logistics-weitu");
        customer_names.put("rohana", "Rayway International Logistics-weitu");
        customer_names.put("rohanawheel", "ROHANA WHEELS");
        customer_names.put("ablcn/rq", "ROHANA WHEELS");
        customer_names.put("rq瑞秋", "RQ");
        customer_names.put("shenzhenrq", "RQ");
        customer_names.put("runyu", "RQ");
        customer_names.put("saunion", "Run Yu");
        customer_names.put("saunion-tesla", "SA UNION");
        customer_names.put("safround", "SA UNION");
        customer_names.put("wuh-宁波顺圆", "Safround Logistics-顺圆");
        customer_names.put("顺圆safround", "Safround Logistics-顺圆");
        customer_names.put("顺圆safround-dfw", "Safround Logistics-顺圆");
        customer_names.put("sfexpress", "Safround Logistics-顺圆");
        customer_names.put("chinainterocean", "SF Express");
        customer_names.put("chinainterocean", "Sinotrans/China Interocean");
        customer_names.put("speedtrucking", "Sinotrans/China Interocean");
        customer_names.put("speedtruck", "Speed Trucking / T-speed Inc");
        customer_names.put("speedtrucking", "Speed Trucking / T-speed Inc");
        customer_names.put("speedmark", "Speed Trucking / T-speed Inc");
        customer_names.put("speedship", "Speedmark");
        customer_names.put("storm", "Speedship 快捷舰");
        customer_names.put("sunnywing", "Storm Logistics");
        customer_names.put("sunnywing", "Sunny Wing/Xu Ying");
        customer_names.put("sunpower", "Sunny Wing/Xu Ying");
        customer_names.put("superlinking", "Sunpower");
        customer_names.put("super-linking", "Super Linking Logistics");
        customer_names.put("superfylg", "Super Linking Logistics");
        customer_names.put("tempest", "Superflyg");
        customer_names.put("tempest", "Tempest");
        customer_names.put("titan", "Tempest");
        customer_names.put("titan", "Titan");
        customer_names.put("tius", "Titan");
        customer_names.put("topline", "Tius Supply Chain Inc");
        customer_names.put("transpacific", "Top Line");
        customer_names.put("transpacific", "Transpacific");
        customer_names.put("twtools", "Transpacific");
        customer_names.put("uqi", "TW TOOLS");
        customer_names.put("viclogistics", "UQI");
        customer_names.put("viclogistics", "VIC Logistics");
        customer_names.put("focusone", "VIC Logistics");
        customer_names.put("vite", "Vite-EasyGo/Focus One 亿龙达");
        customer_names.put("vite/focusone", "Vite-EasyGo/Focus One 亿龙达");
        customer_names.put("vite-focusone", "Vite-EasyGo/Focus One 亿龙达");
        customer_names.put("voodi", "Vite-EasyGo/Focus One 亿龙达");
        customer_names.put("shenzhenwayota", "Voodi");
        customer_names.put("wayota", "Wayota/Huayangda");
        customer_names.put("weship", "Wayota/Huayangda");
        customer_names.put("weship", "Weship");
        customer_names.put("weship/bwg", "Weship");
        customer_names.put("weship-ntlqp", "Weship");
        customer_names.put("wingocean", "Weship");
        customer_names.put("wingocean", "Wing Ocean");
        customer_names.put("wiserbridge", "Wing Ocean");
        customer_names.put("x-change", "Wiser Bridge");
        customer_names.put("xchange360", "Xchange 360");
        customer_names.put("xinbang", "Xchange 360");
        customer_names.put("xinbang鑫邦", "Xin Bang 鑫邦");
        customer_names.put("xinyue", "Xin Bang 鑫邦");
        customer_names.put("xinyue-dallas", "XinYue 信越");
        customer_names.put("xtd", "XinYue 信越");
        customer_names.put("yiwuhangi", "XinYue 信越");
        customer_names.put("yiwuhangji", "XTD-翔通达");
        customer_names.put("yiwuhangji", "YIWU HANGJI");
    }

    public static void main(String[] args) {
        String inputFilePath = "C:\\Users\\CalvinYuen\\Downloads\\TEST_EXCEL\\EXAMPLE_TESTER.xlsx";
        String outputFilePath = "cleaned_output.xlsx";
    
        try (FileInputStream fis = new FileInputStream(new File(inputFilePath));
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {
    
            Sheet sheet = workbook.getSheetAt(0); // Get the first sheet
    
            // Iterate through each row and cell to clean the data
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row != null) {
                    for (int columnIndex = 0; columnIndex < 9; columnIndex++) {
                        Cell cell = row.getCell(columnIndex);
                        if (cell != null) {
                            cleanCell(cell, sheet, rowIndex);
                        }
                    }
                }
            }
    
            // Save the cleaned data back to the output file
            try (FileOutputStream fos = new FileOutputStream(new File(outputFilePath))) {
                workbook.write(fos);
            }
    
            System.out.println("Data cleaning complete. Cleaned file saved as " + outputFilePath);
    
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    

    private static void cleanCell(Cell cell, Sheet sheet, int rowIndex) {
        int columnIndex = cell.getColumnIndex();
        switch (columnIndex) {
            case 0: // CUSTOMER
                customerCleanCell(cell);
                break;
            case 1: // CNTR#
                containerCleanCell(cell, sheet, rowIndex, columnIndex);
                break;
            case 2: // SIZE
                sizeCleanCell(cell);
                break;
            case 3: // ETA PORT
                cleanDateCell(cell);
                break;
            case 4: // ETA
                cleanDateCell(cell);
                break;
            case 5: // SSL
                if (cell.getCellType() == CellType.STRING) {
                    String cellValue = cell.getStringCellValue().toUpperCase().trim();
                    // Replace "CMA" only if it is the only text in the cell
                    if (cellValue.equals("CMA")) {
                        cellValue = "CMA CGM";
                    }
                    // Remove codes that start with 4 capital letters followed by 7 digits
                    cellValue = cellValue.replaceAll("\\b[A-Z]{4}\\d{7}\\b", "").trim();
                    // Set the updated value back to the cell
                    cell.setCellValue(cellValue);
                }
                break;
            case 6: // Rail
                cleanRailCell(cell);
                break;
            case 7: // APPT
                cleanApptCell(cell, sheet, rowIndex);
                break;
            case 8: // TIME
                cleanTimeCell(cell);
                break;
            default:
                // Add other cases as necessary
                break;
        }
    }

    private static void customerCleanCell(Cell cell) {
        if (cell.getCellType() == CellType.STRING) {
            String originalValue = cell.getStringCellValue();
            String tempValue = originalValue.replaceAll("\\s+", "").toLowerCase();
            if (customer_names.containsKey(tempValue)) {
                cell.setCellValue(customer_names.get(tempValue));
            } else {
                cell.setCellValue(originalValue.trim().replaceAll("\\s+", " "));
            }
        }
    }

    private static void containerCleanCell(Cell cell, Sheet sheet, int rowIndex, int columnIndex) {
        if (cell.getCellType() == CellType.STRING) {
            String cellValue = cell.getStringCellValue().toUpperCase();
            cellValue = cellValue.replaceAll("FTL", "").replaceAll("-", "").replaceAll(",", "").replaceAll(" ", "");
    
            // Use regex to find all valid container codes (either four letters followed by 7 digits or "NAM" followed by 7 digits)
            Pattern pattern = Pattern.compile("([A-Z]{4}\\d{7})|(NAM\\d{7})");
            Matcher matcher = pattern.matcher(cellValue);
            List<String> codes = new ArrayList<>();
    
            while (matcher.find()) {
                codes.add(matcher.group());
            }
    
            if (!codes.isEmpty()) {
                cell.setCellValue(codes.get(0));
                for (int i = 1; i < codes.size(); i++) {
                    Row newRow = sheet.createRow(sheet.getLastRowNum() + 1);
                    copyRow(sheet.getRow(rowIndex), newRow, columnIndex, codes.get(i));
                }
            }
        }
    }
    
    private static void copyRow(Row srcRow, Row destRow, int containerColumnIndex, String containerCode) {
        for (int j = 0; j < srcRow.getLastCellNum(); j++) {
            Cell srcCell = srcRow.getCell(j);
            Cell destCell = destRow.createCell(j);
            if (j == containerColumnIndex) {
                destCell.setCellValue(containerCode);
            } else {
                if (srcCell != null) {
                    if (srcCell.getCellType() == CellType.STRING) {
                        destCell.setCellValue(srcCell.getStringCellValue());
                    } else if (srcCell.getCellType() == CellType.NUMERIC) {
                        destCell.setCellValue(srcCell.getNumericCellValue());
                    }
                }
            }
        }
    }
    
    private static void sizeCleanCell(Cell cell) {
        if (cell.getCellType() == CellType.STRING) {
            String cellValue = cell.getStringCellValue().toUpperCase();
            cellValue = cellValue.replaceAll("\\s+", "").replaceAll("-", "");
            cellValue = cellValue.replaceAll("1\\*|1X", "").replaceAll("2\\*|2X", "");
            cellValue = cellValue.replaceAll("3\\*|3X", "").replaceAll("4\\*|4X", "");
            cellValue = cellValue.replace("DC", "GP").replace("DV", "GP");
            cellValue = cellValue.replace("HC", "HQ").replace("HG", "HQ");
            cellValue = cellValue.replace("20>40", "40HQ").replace("40OS", "40HQ");
            cellValue = cellValue.replace("DS", "GP").replace("FT", "GP");
            cellValue = cellValue.replace("?", "HQ").replace("40HRNOR", "40HQ");

            // Remove codes that start with 4 capital letters followed by 7 digits
            cellValue = cellValue.replaceAll("\\b[A-Z]{4}\\d{7}\\b", "").trim();

            cell.setCellValue(cellValue);
        }
    }

    private static void cleanDateCell(Cell cell) {
        if (cell.getCellType() == CellType.NUMERIC) {
            double cellValue = cell.getNumericCellValue();
            Date date = DateUtil.getJavaDate(cellValue);
            SimpleDateFormat sdf = new SimpleDateFormat("MM/dd");
            cell.setCellValue(sdf.format(date));
        } else if (cell.getCellType() == CellType.STRING) {
            String cellValue = cell.getStringCellValue();
            try {
                double numericValue = Double.parseDouble(cellValue);
                Date date = DateUtil.getJavaDate(numericValue);
                SimpleDateFormat sdf = new SimpleDateFormat("MM/dd");
                cell.setCellValue(sdf.format(date));
            } catch (NumberFormatException e) {
                // Not a number, leave it as MM/dd format
                cell.setCellValue(cellValue);
            }
        }
    }

    private static void cleanRailCell(Cell cell) {
        if (cell.getCellType() == CellType.STRING) {
            String cellValue = cell.getStringCellValue().toUpperCase();
            // Remove codes that start with 4 capital letters followed by 7 digits
            cellValue = cellValue.replaceAll("\\b[A-Z]{4}\\d{7}\\b", "").trim();
            
            // Trim the cell value again after replacement
            cellValue = cellValue.trim();
            
            if (cellValue.contains("BNSF")) {
                cell.setCellValue("BNSF");
            } else if (cellValue.contains("BALTIMORE")) {
                cell.setCellValue("BALTIMORE");
            } else if (cellValue.contains("CONLEY")) {
                cell.setCellValue("BOSTON-CONLEY");
            } else if (cellValue.contains("CHARLESTON")) {
                cell.setCellValue("CHARLESTON");
            } else if (cellValue.contains("CSX")) {
                cell.setCellValue("CSX");
            } else if (cellValue.contains("CN")) {
                cell.setCellValue("CN");
            } else if (cellValue.contains("CP")) {
                if (cellValue.contains("LAX TRACPAC")) {
                    // Do nothing, keep the original value
                } else {
                    cell.setCellValue("CP");
                }
            } else if (cellValue.contains("FL") && cellValue.contains("PORT")) {
                cell.setCellValue("FLORIDA PORT");
            } else if (cellValue.contains("HOUSTON")) {
                cell.setCellValue("HOUSTON");
            } else if (cellValue.contains("JACKSONVILLE")) {
                cell.setCellValue("JACKSONVILLE");
            } else if (cellValue.contains("YUSEN")) {
                cell.setCellValue("LAX");
            } else if (cellValue.contains("LAX")) {
                cell.setCellValue("LAX");
            } else if (cellValue.contains("LBCT")) {
                cell.setCellValue("LBCT");
            } else if (cellValue.contains("LGB")) {
                cell.setCellValue("LGB PORT");
            } else if (cellValue.contains("LONG BEACH")) {
                cell.setCellValue("LONG BEACH PORT");
            } else if (cellValue.contains("MAHER")) {
                cell.setCellValue("MAHER TERMINAL");
            } else if (cellValue.contains("MIAMI")) {
                cell.setCellValue("MIAMI");
            } else if (cellValue.contains("MOBILE")) {
                cell.setCellValue("MOBILE");
            } else if (cellValue.contains("NA")) {
                cell.setCellValue("");
            } else if (cellValue.contains("NEW ORLEANS")) {
                cell.setCellValue("NEW ORLEANS");
            } else if (cellValue.contains("NEW YORK")) {
                cell.setCellValue("NEW YORK");
            } else if (cellValue.contains("NORFOLK")) {
                cell.setCellValue("NORFOLK PORT");
            } else if (cellValue.contains("NORFORK")) {
                cell.setCellValue("NORFOLK PORT");
            } else if (cellValue.contains("UP")) {
                cell.setCellValue("UP");
            } else if (cellValue.contains("NS ")) {
                cell.setCellValue("NS PORT");
            } else if (cellValue.contains("NS-")) {
                cell.setCellValue("NS PORT");
            } else if (cellValue.contains("OAKLAND")) {
                cell.setCellValue("OAKLAND");
            } else if (cellValue.contains("PORT LIBERTY")) {
                cell.setCellValue("PORT LIBERTY");
            } else if (cellValue.contains("PORT OF TAMPA")) {
                cell.setCellValue("PORT OF TAMPA");
            } else if (cellValue.contains("SAV")) {
                cell.setCellValue("SAVANNAH PORT");
            } else if (cellValue.contains("SEAT")) {
                cell.setCellValue("SEATTLE PORT");
            } else if (cellValue.contains("SOUTH FLORIDA")) {
                cell.setCellValue("SOUTH FLORIDA");
            } else if (cellValue.contains("TACOMA")) {
                cell.setCellValue("TACOMA PORT");
            } else if (cellValue.contains("TAMPA")) {
                cell.setCellValue("PORT OF TAMPA");
            } else if (cellValue.contains("UP ")) {
                cell.setCellValue("UP");
            } else if (cellValue.contains("UP-")) {
                cell.setCellValue("UP");
            } else if (cellValue.contains("WIL")) {
                cell.setCellValue("WILMINGTON PORT");
            } else if (cellValue.contains("FTL")) {
                cell.setCellValue("");
            } else {
                cell.setCellValue(cellValue);
            }
        }
    }    

    private static void cleanApptCell(Cell cell, Sheet sheet, int rowIndex) {
        if (cell.getCellType() == CellType.NUMERIC) {
            Date date = DateUtil.getJavaDate(cell.getNumericCellValue());
            SimpleDateFormat sdf = new SimpleDateFormat("MM-dd");
            cell.setCellValue(sdf.format(date));
        } else if (cell.getCellType() == CellType.STRING) {
            String cellValue = cell.getStringCellValue().toUpperCase();
    
            // Extract and remove times (including time zones) if they exist
            Pattern timePattern = Pattern.compile("\\b\\d{1,2}:\\d{2}(?:\\s?(AM|PM))?(?:\\s?[A-Z]+)?\\b");
            Matcher timeMatcher = timePattern.matcher(cellValue);
            List<String> extractedTimes = new ArrayList<>();
            while (timeMatcher.find()) {
                extractedTimes.add(timeMatcher.group());
            }
            cellValue = timeMatcher.replaceAll("").trim(); // Remove times from cell value
    
            // Remove specific phrases and unwanted characters
            String[] phrasesToRemove = {
                "EXAM HOLD AT LAX", "LAX EXAM", "CET EXAM", "EXAM HOLD", "GUARDIAN LOGISTICS SOLUTION CHARLOTTE, NC 28208", 
                "最少提前3天预约", "@", "AT", "DROP", "ABL", "DOCK#", "2305008569 HOLD"
            };
            
            for (String phrase : phrasesToRemove) {
                cellValue = cellValue.replace(phrase, "").trim();
            }
    
            // Check if the cell contains any numbers
            if (!cellValue.matches(".*\\d.*")) {
                cell.setCellValue("");  // Remove the text if no numbers are present
                return;
            }
    
            // Define the list to hold extracted dates
            List<String> dates = new ArrayList<>();
    
            // Define regex patterns for dates
            Pattern datePattern1 = Pattern.compile("\\b\\d{1,2}-\\d{1,2}\\b");  // Matches MM-dd
            Pattern datePattern2 = Pattern.compile("\\b\\d{1,2}/\\d{1,2}\\b");  // Matches MM/dd
            Pattern datePattern3 = Pattern.compile("\\b\\d{1,2}/\\d{1,2}/\\d{4}\\b");  // Matches MM/dd/yyyy
            Pattern datePattern4 = Pattern.compile("\\b\\d{4}-\\d{2}-\\d{2}\\b");  // Matches yyyy-MM-dd
    
            // Extract and convert numbers to dates if possible
            Matcher matcher1 = datePattern1.matcher(cellValue);
            while (matcher1.find()) {
                dates.add(matcher1.group());
            }
            
            Matcher matcher2 = datePattern2.matcher(cellValue);
            while (matcher2.find()) {
                dates.add(matcher2.group());
            }
    
            Matcher matcher3 = datePattern3.matcher(cellValue);
            while (matcher3.find()) {
                dates.add(matcher3.group());
            }
    
            Matcher matcher4 = datePattern4.matcher(cellValue);
            while (matcher4.find()) {
                dates.add(matcher4.group());
            }
    
            // If dates were found, format them to MM-dd and join them into a single cell value
            if (!dates.isEmpty()) {
                List<String> formattedDates = new ArrayList<>();
                SimpleDateFormat outputFormat = new SimpleDateFormat("MM-dd");
                for (String dateStr : dates) {
                    try {
                        Date date;
                        if (dateStr.contains("/")) {
                            SimpleDateFormat inputFormat = new SimpleDateFormat(dateStr.length() == 5 ? "MM/dd" : "MM/dd/yyyy");
                            date = inputFormat.parse(dateStr);
                        } else {
                            SimpleDateFormat inputFormat = new SimpleDateFormat(dateStr.length() == 5 ? "MM-dd" : "yyyy-MM-dd");
                            date = inputFormat.parse(dateStr);
                        }
                        formattedDates.add(outputFormat.format(date));
                    } catch (Exception e) {
                        // Ignore invalid date formats
                    }
                }
                cell.setCellValue(String.join(" ", formattedDates));
            } else {
                cell.setCellValue(cellValue);
            }
    
            // If times were extracted, add them to the TIME column (index 8)
            if (!extractedTimes.isEmpty()) {
                Row row = sheet.getRow(rowIndex);
                Cell timeCell = row.getCell(8);
                if (timeCell == null) {
                    timeCell = row.createCell(8);
                }
                if (timeCell.getCellType() == CellType.BLANK) {
                    timeCell.setCellValue(String.join(" ", extractedTimes));
                } else {
                    String existingTimes = timeCell.getStringCellValue();
                    timeCell.setCellValue(existingTimes + " " + String.join(" ", extractedTimes));
                }
            }
        }
    }    
    
    private static void cleanTimeCell(Cell cell) {
        if (cell.getCellType() == CellType.NUMERIC) {
            double numericValue = cell.getNumericCellValue();
            Date date = DateUtil.getJavaDate(numericValue);
            if (date != null) {
                SimpleDateFormat sdf = new SimpleDateFormat("hh:mm a");
                cell.setCellValue(sdf.format(date));
            } else {
                System.out.println("Failed to parse date from numeric value: " + numericValue);
            }
        } else if (cell.getCellType() == CellType.STRING) {
            String cellValue = cell.getStringCellValue().toUpperCase().trim();
    
            // Remove "DROP" from the cell content
            cellValue = cellValue.replaceAll("DROP", "").trim();
    
            // Remove codes that start with 4 capital letters followed by 7 digits
            cellValue = cellValue.replaceAll("\\b[A-Z]{4}\\d{7}\\b", "").trim();
    
            // Replace "CMA" only if it is the only text in the cell
            if (cellValue.equals("CMA")) {
                cellValue = "CMA CGM";
            }
    
            // Define regex pattern to match time formats (e.g., 10:30 AM, 14:30, etc.)
            Pattern timePattern = Pattern.compile("\\b\\d{1,2}:\\d{2}(?:\\s?(AM|PM))?\\b");
            Matcher matcher = timePattern.matcher(cellValue);
    
            List<String> times = new ArrayList<>();
    
            while (matcher.find()) {
                String time = matcher.group();
                try {
                    SimpleDateFormat inputFormat = new SimpleDateFormat("hh:mm a");
                    SimpleDateFormat outputFormat = new SimpleDateFormat("hh:mm a");
                    Date date = inputFormat.parse(time);
                    times.add(outputFormat.format(date));
                } catch (ParseException e) {
                    times.add(time); // Keep the original time if parsing fails
                }
            }
    
            // Check for times like "1PM" or "11AM"
            Pattern simpleTimePattern = Pattern.compile("\\b\\d{1,2}(AM|PM)\\b");
            Matcher simpleTimeMatcher = simpleTimePattern.matcher(cellValue);
            while (simpleTimeMatcher.find()) {
                String time = simpleTimeMatcher.group();
                try {
                    SimpleDateFormat inputFormat = new SimpleDateFormat("hha");
                    SimpleDateFormat outputFormat = new SimpleDateFormat("hh:mm a");
                    Date date = inputFormat.parse(time);
                    times.add(outputFormat.format(date));
                } catch (ParseException e) {
                    times.add(time); // Keep the original time if parsing fails
                }
            }
    
            // Check for numbers that could be converted to time
            String[] parts = cellValue.split("\\s+");
            for (String part : parts) {
                try {
                    double numericValue = Double.parseDouble(part);
                    Date date = DateUtil.getJavaDate(numericValue);
                    if (date != null) {
                        SimpleDateFormat sdf = new SimpleDateFormat("hh:mm a");
                        times.add(sdf.format(date));
                    } else {
                        System.out.println("Failed to parse date from numeric part: " + part);
                    }
                } catch (NumberFormatException e) {
                    // Not a valid number, ignore
                }
            }
    
            if (!times.isEmpty()) {
                cell.setCellValue(String.join(" ", times));
            } else {
                cell.setCellValue(""); // Clear cell if no valid time formats are found
            }
        }
    }     
}
    