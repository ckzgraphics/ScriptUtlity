package org.base.handler;

import java.io.*;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class ExcelHandler {

    private static final Logger LOG = LoggerFactory.getLogger(ExcelHandler.class);

    private Workbook workbook = null;
    private Sheet sheet = null;
    private Row row = null;
    private Row rowZero = null;
    private Cell cell = null;

    private int totalRowCount = 0;
    private int totalColumnCount = 0;
    private Map<Integer, String> columnHeading = null;

    private int totalNumberOfSheets = 0;

    private FileInputStream inputStream = null;

    /**
     * To read the excel file call method: readExcelFile
     */
    public ExcelHandler() {
    }

    /**
     * Reads excel workbook (call method: setSheet to read sheet)
     *
     * @param filePath  Absolute path to the file directory
     * @param fileName  Excel file name with extension
     */
    public ExcelHandler(String filePath, String fileName) {
        setWorkbook(filePath, fileName);
    }

    /**
     * Reads excel file (Do not call method: readExcelFile explicitly)
     *
     * @param filePath  Absolute path to the file directory
     * @param fileName  Excel file name with extension
     * @param sheetName Excel sheet name
     */
    public ExcelHandler(String filePath, String fileName, String sheetName) {
        readExcelFile(filePath, fileName, sheetName);
    }

    /**
     * Reads excel workbook
     *
     * @param filePath Absolute path to the file directory
     * @param fileName Excel file name with extension
     */
    public boolean setWorkbook(String filePath, String fileName) {
        LOG.info("Reading file: {} from path: {}", fileName, filePath);
        boolean isComplete = false;
        String fileExtension = null;
        try {
            fileExtension = FilenameUtils.getExtension(fileName);
            System.out.println("Reading File: " + filePath + fileName);
            LOG.info("File extension: {}", fileExtension);
            inputStream = new FileInputStream(new File(filePath.concat(File.separator).concat(fileName)));
            this.workbook = WorkbookFactory.create(inputStream);
            isComplete = true;
            totalNumberOfSheets = workbook.getNumberOfSheets();
        } catch (NullPointerException e) {
            System.out.println("Error while reading excel file " +  fileName + " :: " + e.getMessage());
            e.printStackTrace();
        } catch (IOException e) {
            System.out.println("IO error while reading excel file " +  fileName + " :: " + e.getMessage());
            e.printStackTrace();
        } catch (Exception e) {
            System.out.println("Error while reading excel file " +  fileName + " :: " + e.getMessage());
            e.printStackTrace();
        }
        return isComplete;
    }

    public void closeWorkbook(){
        try {
            if(Objects.nonNull(this.workbook)){
                workbook.close();
            }

        } catch(Exception e){
            System.out.println("Error while closing workbook");
            e.printStackTrace();
        }
    }

    /**
     *
     * @param sheetName
     */
    public boolean setSheet(String sheetName) {
        boolean isComplete = false;
        this.sheet = workbook.getSheet(sheetName);
        System.out.println("Reading sheet: " + sheetName);
        LOG.info("Reading sheet: {}", sheetName);
        if(this.sheet != null) {
            // IN POC ROW 0 HAS NUM VALUE
            // AND ROW 1 IS COLS VALUE
            this.rowZero = this.sheet.getRow(1);
            this.totalRowCount = getActualRowCount();
            this.totalColumnCount = getActualColumnCount();
            this.columnHeading = getColumnHeading(rowZero);
            /*
             * FOLLOWING DEFAULT METHOD DOES NOT WORK AS EXPECTED FOR EXCEL FILE CONVERTED
             * FROM GOOGLE SHEET HENCE CREATED CUSTOM METHOD getActualColumnCount()
             */
            // this.totalColumnCount = sheet.getRow(0).getLastCellNum();
            System.out.println("Rows:  "+this.totalRowCount+"| Columns: "+this.totalColumnCount);
            LOG.info("Rows: {} | Columns: {}", this.totalRowCount, this.totalColumnCount);
            isComplete = true;
        } else {
            isComplete = true;
        }
        return isComplete;
    }

    /**
     *
     * @param sheetNumber
     */
    public void setSheet(int sheetNumber) {
        this.sheet = workbook.getSheetAt(sheetNumber);
        LOG.info("Reading sheet: {}", sheetNumber);
        if (sheet != null) {
            this.rowZero = this.sheet.getRow(0);
            this.totalRowCount = getActualRowCount();
            this.totalColumnCount = getActualColumnCount();
            this.columnHeading = getColumnHeading(rowZero);
            /*
             * FOLLOWING DEFAULT METHOD DOES NOT WORK AS EXPECTED FOR EXCEL FILE CONVERTED
             * FROM GOOGLE SHEET HENCE CREATED CUSTOM METHOD getActualColumnCount()
             */
            // this.totalColumnCount = sheet.getRow(0).getLastCellNum();
            LOG.info("Rows: {} | Columns: {}", this.totalRowCount, this.totalColumnCount);
        }
    }

    /**
     * Reads excel workbook and sheet
     *
     * @param filePath  Absolute path to the file directory
     * @param fileName  Excel file name with extension
     * @param sheetName Excel sheet name
     */
    public void readExcelFile(String filePath, String fileName, String sheetName) {
        try {
            setWorkbook(filePath, fileName);
        } catch (Exception e) {
            LOG.error("Error while reading excel file {} :: {}", fileName, e.getMessage());
        } finally {
            if (this.workbook != null) {
                setSheet(sheetName);
            }
        }
    }

    /**
     *
     * @return Two row count
     */
    public int getRowCount() {
        return totalRowCount;
    }

    /**
     *
     * @return Total column count
     */
    public int getColumnCount() {
        return totalColumnCount;
    }

    /**
     *
     * @return Column Count
     */
    private int getActualColumnCount() {
        Cell cell = null;
        int rowZeroColumnCount = rowZero.getLastCellNum();
        int colIndex = 0;
        for (; colIndex < rowZeroColumnCount; colIndex++) {
            cell = rowZero.getCell(colIndex);
            if (isCellBlankOrNull(cell) == true) {
                break;
            }
        }
        LOG.info("Columns found: {}, Physical columns found: {}", rowZeroColumnCount, colIndex);
        return colIndex;
    }

    /**
     *
     * @return Row Count
     */
    private int getActualRowCount() {
        Cell cell = null;
        int rowCount = this.sheet.getPhysicalNumberOfRows();
        int rowIndex = 0;
        for (; rowIndex < rowCount; rowIndex++) {
            cell = this.sheet.getRow(rowIndex).getCell(0);
            if (isCellBlankOrNull(cell) == true) {
                break;
            }
        }
        LOG.info("Rows found: {}, Physical rows found: {}", rowCount, rowIndex - 1);
        return --rowIndex;
    }

    /**
     *
     * @param rowZero Row object containing Zeroth row
     * @return Mapping of Column Index and Column Name
     */
    private Map<Integer, String> getColumnHeading(Row rowZero) {
        String colName = null;
        Map<Integer, String> columnHeading = new LinkedHashMap<Integer, String>();
        for (int colIndex = 0; colIndex < totalColumnCount; colIndex++) {
            colName = rowZero.getCell(colIndex).getStringCellValue();
            columnHeading.put(colIndex, colName);
        }
        return columnHeading;
    }

    /**
     *
     * @param cell
     * @return String value of the cell
     */
    public String getCellStringValue(Cell cell) {
        CellType cellType = null;
        String cellValue = "";
        if(Objects.nonNull(cell)){
            cellType = cell.getCellType();
            if (cellType != null) {
                switch (cellType) {
                    case STRING:
                        cellValue = cell.getStringCellValue();
                        break;
                    case NUMERIC:
                        cellValue = Double.toString((cell.getNumericCellValue()));
                        break;
                    case BOOLEAN:
                        cellValue = Boolean.toString(cell.getBooleanCellValue());
                        break;
                    case _NONE:
                        cellValue = "";
                        break;
                    case BLANK:
                        cellValue = "";
                        break;
                    case ERROR:
                        cellValue = Byte.toString(cell.getErrorCellValue());
                        break;
                    case FORMULA:
//                    cellValue = Integer.toString(((int) cell.getNumericCellValue()));
                        cellValue = "NaN";
                        break;
                    default:
                        cellValue = "";
                }
            }
        } else {
            System.out.println("NC");
        }
        return cellValue;
    }

    /**
     *
     * @param cell
     * @return If NULL/BLANK: true | If value found: false
     */
    public boolean isCellBlankOrNull(Cell cell) {
        if ((cell.getCellType() == CellType.BLANK) || cell.getCellType() == CellType._NONE)
            return true;
        else
            return false;
    }

    /**
     * Prints excel data
     */
    public void printSheetData() {
        boolean lastRowFlag = false;
        String tempStr = "";
        for (int rowIndex = 0; rowIndex <= this.totalRowCount; rowIndex++) {
            tempStr += "\n\nROW " + rowIndex;
            this.row = this.sheet.getRow(rowIndex);
            for (int colIndex = 0; colIndex < this.totalColumnCount; colIndex++) {
                this.cell = this.row.getCell(colIndex);
                if (colIndex == 0) {
                    if (isCellBlankOrNull(this.cell) == true) {
                        lastRowFlag = true;
                        break;
                    }
                }
                tempStr += "," + getCellStringValue(this.cell);
            }
            if (lastRowFlag)
                break;
        }
        LOG.info("Printing the excel data..\n {}", tempStr);
    }

    public void printSheetData(String sheetName) {
        boolean lastRowFlag = false;
        String tempStr = "";
        setSheet(sheetName);
        for (int rowIndex = 0; rowIndex <= this.totalRowCount; rowIndex++) {
            tempStr += "\n\nROW " + rowIndex;
            this.row = this.sheet.getRow(rowIndex);
            for (int colIndex = 0; colIndex < this.totalColumnCount; colIndex++) {
                this.cell = this.row.getCell(colIndex);
                if (colIndex == 0) {
                    if (isCellBlankOrNull(this.cell) == true) {
                        lastRowFlag = true;
                        break;
                    }
                }
                System.out.println("["+rowIndex+","+colIndex+"]");
                tempStr += "||" + getCellStringValue(this.cell);
            }
            if (lastRowFlag)
                break;
        }
        System.out.println("Printing the excel data..\n " + tempStr);
        LOG.info("Printing the excel data..\n {}", tempStr);
    }

    /**
     *
     * @param columnIndex
     * @return Column Name
     */
    public String getColumnName(int columnIndex) {
        return columnHeading.get(columnIndex);
    }

    /**
     *
     * @param columnName
     * @return Column Index
     */
    public int getColumnIndex(String columnName) {
        String tempStr = "";
        int colIndex = 0;
        for (; colIndex < totalColumnCount; colIndex++) {
            tempStr = rowZero.getCell(colIndex).getStringCellValue();
            if (Objects.nonNull(tempStr) && tempStr.equals(columnName)) {
                tempStr = rowZero.getCell(colIndex).getStringCellValue();
                break;
            }
        }
        return colIndex;
    }

    /**
     *
     * @return Excel data
     */
    public List<Map<String, String>> getSheetData_String() {
        Cell cell = null;
        Row row = null;
        List<Map<String, String>> data = new ArrayList<Map<String, String>>();

        for (int rowIndex = 1; rowIndex <= this.totalRowCount; rowIndex++) {
            row = this.sheet.getRow(rowIndex);
            Map<String, String> rowData = new LinkedHashMap<String, String>();

            for (int colIndex = 0; colIndex < this.totalColumnCount; colIndex++) {
                cell = row.getCell(colIndex);
                rowData.put(getColumnName(colIndex), getCellStringValue(cell));
            } // COL END
            data.add(rowData);
        }
        return data;
    }

    /**
     * Reads sheet and
     *
     * @return
     */
    public List<Map<String, Object>> getSheetData() {
        Cell cell = null;
        Row row = null;
        List<Map<String, Object>> data = new ArrayList<Map<String, Object>>();

        for (int rowIndex = 1; rowIndex <= this.totalRowCount; rowIndex++) {
            row = this.sheet.getRow(rowIndex);
            Map<String, Object> rowData = new LinkedHashMap<String, Object>();

            for (int colIndex = 0; colIndex < this.totalColumnCount; colIndex++) {
                cell = row.getCell(colIndex);
                rowData.put(getColumnName(colIndex), cell);
            }
            data.add(rowData);
        }
        return data;
    }

    public int getNumberOfSheets(){return totalNumberOfSheets;}

    public List<String> getSheetNameList(){

        String sheetNamesPrint = "";
        List<String> sheetNames = new ArrayList<>();
        System.out.print("\nSheetNames => ");
        for(int i=0; i<totalNumberOfSheets; i++) {
            sheetNamesPrint += i + ":" + workbook.getSheetAt(i).getSheetName() + "|";
            sheetNames.add(workbook.getSheetName(i));
        }
        System.out.print(sheetNamesPrint + "\n\n");
        return sheetNames;
    }

    public void addRowData(String sheetName, List<String> data){

        setSheet(sheetName);
        Row row = sheet.createRow(this.totalRowCount + 1);

        if(this.totalColumnCount == data.size()){
            for (int colIndex = 0; colIndex < this.totalColumnCount; colIndex++) {
                row.createCell(colIndex).setCellValue(data.get(colIndex));
            }
        } else {
            System.out.println("Count of data to be added: "+data.size());
            // SAVE TO FIRST 13 COLs ONLY
            for (int colIndex = 0; colIndex < 13; colIndex++) {
                row.createCell(colIndex).setCellValue(data.get(colIndex));
            }
        }
    }

    public  void saveFile(String filePath, String fileName){
        OutputStream fileOut = null;
        try {
            if(Objects.nonNull(inputStream))
                inputStream.close();
            fileOut = new FileOutputStream(filePath + fileName);
            workbook.write(fileOut);
        } catch(Exception e){
            System.out.println("Error while saving the file");
            e.printStackTrace();
        } finally {
             if(Objects.nonNull(fileOut)){
                 try {
                     fileOut.flush();
                     fileOut.close();
                 } catch (Exception e){e.printStackTrace();}
             }
        }
    }


}
