import com.fasterxml.jackson.annotation.JsonFormat;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xddf.usermodel.*;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xssf.usermodel.*;
import org.json.simple.parser.JSONParser;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

public class LevelingDataParser {

    private static final String levelingDataFile = "C:\\Program Files (x86)\\World of Warcraft\\_beta_\\WTF\\Account\\975281#4\\SavedVariables\\PoliLevelingUtil.lua";
    private static final String destFileLocation = "C:\\Users\\Adam\\Desktop\\WoWLevelingReports\\";

    private static final Map<Integer, Integer> xpPerLevel;
    static {
        Map<Integer, Integer> map = new HashMap<>();
        map.put(50, 178215);
        map.put(51, 205775);
        map.put(52, 242790);
        map.put(53, 272185);
        map.put(54, 315745);
        map.put(55, 355500);
        map.put(56, 397520);
        map.put(57, 441860);
        map.put(58, 488565);
        map.put(59, 515400);
        xpPerLevel = Collections.unmodifiableMap(map);
    }

    public static void main(String[] args) throws IOException, org.json.simple.parser.ParseException, ParseException {
        String levelingDataJsonString = getLevelingDataJsonString(levelingDataFile);
        ObjectMapper mapper = new ObjectMapper();
        mapper.setTimeZone(TimeZone.getDefault());
        List<LevelingSession> sessions = mapper.readValue(levelingDataJsonString, new TypeReference<>(){});
        sessionsToSpreadsheets(sessions.toArray(new LevelingSession[0]), destFileLocation);
    }

    private static String getLevelingDataJsonString(String filename) throws IOException, org.json.simple.parser.ParseException {
        Path file = Path.of(filename);
        String fileDataString = Files.readString(file);
        String result = fileDataString.substring(fileDataString.indexOf("[{"), fileDataString.indexOf("\"\r\n"));
        result = result.replaceAll("\\\\", "");
        return new JSONParser().parse(result).toString();
    }

    private static void sessionsToSpreadsheets(LevelingSession[] sessions, String fileLocation) throws IOException  {

        for (int i=0; i<sessions.length; i++) {
            String wbName = "LevelingData_" + new SimpleDateFormat("MMddyy_HHmmss").format(sessions[i].loginTime) + ".xlsx";
            XSSFWorkbook wb = new XSSFWorkbook();
            createSummarySheet(sessions[i], wb);
            Recording[] recordings = sessions[i].recordings;
            if (recordings.length == 0) {
                return;
            }
            for (int j=0; j<recordings.length; j++) {
                Recording recording = recordings[j];
                XSSFSheet sheet = wb.createSheet(createSheetName(recording, wb, j));

                Row row = sheet.createRow(0);
                row.createCell(0).setCellValue(0);
                row.createCell(1).setCellValue(recording.characterLevel);
                for (int k=1; k<=recording.snapshots.length; k++) {
                    row = sheet.createRow(k);
                    row.createCell(0).setCellValue(k * 5);
                    row.createCell(1).setCellValue(recording.snapshots[k-1]);
                }
                createChartOnSheet(sheet, recording.snapshots.length, 1, 0, 1, "Time (seconds)", "Level");

                row = sheet.getRow(0);
                row.createCell(2).setCellValue(0);
                row.createCell(3).setCellValue(0);
                int k;
                for (k=11; k<recording.snapshots.length; k+=12) {
                    row = sheet.getRow((k + 1) / 12);
                    row.createCell(2).setCellValue((k + 1) / 12);
                    if (k == 11) {
                        row.createCell(3).setCellValue(recording.snapshots[11] - recording.characterLevel);
                    } else {
                        row.createCell(3).setCellValue(recording.snapshots[k] - recording.snapshots[k - 12]);
                    }
                }
                if (k != 11) {
                    if (k-recording.snapshots.length != 11) {
                        row = sheet.getRow((k + 1) / 12);
                        row.createCell(2).setCellValue((k/12*12+11-(k-recording.snapshots.length)) / 12.0);
                        row.createCell(3).setCellValue(recording.snapshots[recording.snapshots.length-1] - recording.snapshots[recording.snapshots.length-12+(k-recording.snapshots.length)]);
                    }
                    createChartOnSheet(sheet, recording.snapshots.length % 12 == 0 ? recording.snapshots.length / 12 : recording.snapshots.length / 12 + 1, 2, 2, 3, "Time (minutes)", "Level Difference");
                }
            }
            File excelFile = new File(fileLocation + wbName);
            excelFile.getParentFile().mkdirs();
            excelFile.createNewFile();
            FileOutputStream fileOut = new FileOutputStream(excelFile);
            wb.write(fileOut);
            fileOut.close();
        }
    }

    private static String createSheetName(Recording recording, XSSFWorkbook wb, int recordingNumber) {
        String sheetName = !recording.getName().isEmpty() ? recording.getName() : "Sheet " + (recordingNumber+2);
        sheetName = WorkbookUtil.createSafeSheetName(sheetName);
        int copyNumber = 1;
        while (wb.getSheet(sheetName) != null) {
            if (sheetName.indexOf("(\\d+)$") != -1) {
                int index1 = sheetName.lastIndexOf("(");
                sheetName = sheetName.substring(0, index1+1) + copyNumber + ")";
            } else {
                sheetName = sheetName + " (" + copyNumber + ")";
            }
            copyNumber++;
        }
        return sheetName;
    }

    private static void createSummarySheet(LevelingSession session, XSSFWorkbook wb) {
        XSSFSheet sheet = wb.createSheet("Summary");
        Row row = sheet.createRow(1);
        row.createCell(1).setCellValue("Login:");
        row.createCell(2).setCellValue(new SimpleDateFormat("MM/dd/yy HH:mm:ss").format(session.loginTime));
        row = sheet.createRow(2);
        row.createCell(1).setCellValue("Logout:");
        if (session.logoutTime != null) {
            row.createCell(2).setCellValue(new SimpleDateFormat("MM/dd/yy HH:mm:ss").format(session.logoutTime));
        } else {
            row.createCell(2).setCellValue("In Progress");
        }
        row = sheet.createRow(4);
        row.createCell(1).setCellValue("Recordings");
        for (int i=0; i<session.recordings.length; i++) {
            Recording recording = session.recordings[i];
            row = sheet.createRow(6 + i * 5);
            row.createCell(2).setCellValue("Name:");
            row.createCell(3).setCellValue(recording.name);
            row.createCell(6).setCellValue("XP:");
            int xpGain = getXPGain(recording);
            row.createCell(7).setCellValue(xpGain);
            row = sheet.createRow(7 + i * 5);
            row.createCell(2).setCellValue("Duration(minutes):");
            double duration;
            if (recording.stopTime != null) {
                duration = (recording.stopTime.getTime() - recording.startTime.getTime()) / 1000.0 / 60.0;
            } else {
                duration = recording.snapshots.length * 5 / 60.0;
            }
            row.createCell(3).setCellValue(duration);
            row.createCell(6).setCellValue("%XP:");
            double percentXPGain = (recording.snapshots[recording.snapshots.length-1] - recording.characterLevel) * 100;
            row.createCell(7).setCellValue(percentXPGain);
            row = sheet.createRow(8 + i * 5);
            row.createCell(2).setCellValue("Start Level:");
            row.createCell(3).setCellValue(recording.characterLevel);
            row.createCell(6).setCellValue("XP/minute:");
            row.createCell(7).setCellValue(xpGain / duration);
            row = sheet.createRow(9 + i * 5);
            row.createCell(2).setCellValue("Stop Level:");
            row.createCell(3).setCellValue(recording.snapshots[recording.snapshots.length-1]);
            row.createCell(6).setCellValue("%XP/minute:");
            row.createCell(7).setCellValue(percentXPGain / duration);
        }
        sheet.autoSizeColumn(1);
        sheet.autoSizeColumn(2);
        sheet.autoSizeColumn(3);
        sheet.autoSizeColumn(6);
        sheet.autoSizeColumn(7);
    }

    private static int getXPGain(Recording recording) {
        double startLevelXP = recording.characterLevel;
        double stopLevelXP = recording.snapshots[recording.snapshots.length-1];
        int startLevel = (int)Math.floor(startLevelXP);
        int stopLevel = (int)Math.floor(stopLevelXP);
        if (startLevel == stopLevel) {
            return (int)((stopLevelXP - startLevel) * xpPerLevel.get(startLevel));
        } else {
            int sum = xpPerLevel.get(startLevel) - (int)((startLevelXP - startLevel) * xpPerLevel.get(startLevel));
            for (int i=startLevel+1; i<=stopLevel; i++) {
                if (i == stopLevel) {
                    sum += (int)((stopLevelXP - stopLevel) * xpPerLevel.get(i));
                } else {
                    sum += xpPerLevel.get(i);
                }
            }
            return sum;
        }
    }

    private static void createChartOnSheet(XSSFSheet sheet, int rows, int chartNumber, int xCol, int yCol, String bottomTitle, String leftTitle) {
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 4, 1 + (chartNumber-1) * 56, 38, 55 + (chartNumber-1) * 56);

        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText(sheet.getSheetName());
        XDDFValueAxis bottomAxis = chart.createValueAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.BOTTOM);
        XDDFValueAxis leftAxis = chart.createValueAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.LEFT);

        bottomAxis.setTitle(bottomTitle);
        leftAxis.setTitle(leftTitle);
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

        XDDFNumericalDataSource<Double> xs = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(0, rows, xCol, xCol));
        XDDFNumericalDataSource<Double> ys = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(0, rows, yCol, yCol));
        XDDFScatterChartData data = (XDDFScatterChartData) chart.createData(ChartTypes.SCATTER, bottomAxis, leftAxis);
        XDDFScatterChartData.Series series1 = (XDDFScatterChartData.Series) data.addSeries(xs, ys);
        series1.setSmooth(false);
        XDDFSolidFillProperties lineFill = new XDDFSolidFillProperties();
        lineFill.setColor(XDDFColor.from(SchemeColor.ACCENT_1));
        XDDFLineProperties lineProperties = new XDDFLineProperties();
        lineProperties.setFillProperties(lineFill);
        series1.setLineProperties(lineProperties);

        chart.plot(data);
    }

    static class LevelingSession {
        @JsonFormat (shape = JsonFormat.Shape.STRING, pattern = "MM/dd/yy HH:mm:ss" )
        Date loginTime;
        @JsonFormat (shape = JsonFormat.Shape.STRING, pattern = "MM/dd/yy HH:mm:ss")
        Date logoutTime;
        Recording[] recordings;

        public Date getLoginTime() {
            return loginTime;
        }

        public void setLoginTime(Date loginTime) {
            this.loginTime = loginTime;
        }

        public Date getLogoutTime() {
            return logoutTime;
        }

        public void setLogoutTime(Date logoutTime) {
            this.logoutTime = logoutTime;
        }

        public Recording[] getRecordings() {
            return recordings;
        }

        public void setRecordings(Recording[] recordings) {
            this.recordings = recordings;
        }
    }

    static class Recording {
        @JsonFormat (shape = JsonFormat.Shape.STRING, pattern = "MM/dd/yy HH:mm:ss")
        Date startTime;
        @JsonFormat (shape = JsonFormat.Shape.STRING, pattern = "MM/dd/yy HH:mm:ss")
        Date stopTime;
        String name;
        int startGetTime;
        double characterLevel;
        Double[] snapshots;

        public Date getStartTime() {
            return startTime;
        }

        public void setStartTime(Date startTime) {
            this.startTime = startTime;
        }

        public Date getStopTime() {
            return stopTime;
        }

        public void setStopTime(Date stopTime) {
            this.stopTime = stopTime;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public int getStartGetTime() {
            return startGetTime;
        }

        public void setStartGetTime(int startGetTime) {
            this.startGetTime = startGetTime;
        }

        public double getCharacterLevel() {
            return characterLevel;
        }

        public void setCharacterLevel(double characterLevel) {
            this.characterLevel = characterLevel;
        }

        public Double[] getSnapshots() {
            return snapshots;
        }

        public void setSnapshots(Double[] snapshots) {
            this.snapshots = snapshots;
        }
    }
}
