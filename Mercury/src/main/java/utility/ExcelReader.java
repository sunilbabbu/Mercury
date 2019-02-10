
package utility;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;




/**
 * This class is to handle multiple reading operations of content from Excel workbook.
 * @author Sujoy
 *
 */
public class ExcelReader {

    /** Path to the spreadsheet. */
    private String path;
    /** Stream to read the spreadsheet. */
    private FileInputStream fis = null;
    /** Access to the spreadsheet. */
    private XSSFWorkbook workbook = null;
    /** Work sheet in the current file. */ 
    private XSSFSheet sheet = null;
    /** Row within the current sheet. */
    private XSSFRow row = null;
    /** Cell within the current sheet. */
    private XSSFCell cell = null;
    /** Stream to write the spreadsheet. */
    private FileOutputStream fileOut = null;

    /** Log4j Logger. */
    private static final Logger LOG = LogManager.getLogger(ExcelReader.class);

    /**
     * ExcelReader(String path) constructor accepts only the path of the Excel
     * sheet.
     *
     * @author Sujoy
     * @param path The name of and path to the spreadsheet
     *
     **/
    public ExcelReader(final String path) {
        if (path != null) {
            this.setPath(path);

            if (this.getPath() != null) {
                try {
                    setFis(new FileInputStream(this.getPath()));

                    if (this.getFis() != null) {
                        XSSFWorkbook xssfWB = new XSSFWorkbook(getFis());
                        if (xssfWB != null) {
                            setWorkbook(xssfWB);
                        }
                    }
                } catch (FileNotFoundException fe) {
                    LOG.warn("Unable to open spreadsheet " + fe.getMessage());
                } catch (Exception e) {
                    LOG.error("Problem creating stream for path \"" + path + "\"", e);
                } finally {
                    if (this.getFis() != null) {
                        try {
                            getFis().close();
                        } catch (IOException e) {
                            LOG.error(e.getStackTrace());
                        }
                    }
                }
            }
        } else {
            LOG.warn("Unable to open spreadsheet as path is null");
        }
    }

    /**
     * ExcelReader(String path, String SheetName) constructor accepts the path
     * of the Excel sheet with its Name and the workSheetName.
     *
     * @author Sujoy
     * @param path - Name & path of the Excel spreadsheet
     * @param sheetName - Name of the workSheet
     **/

    public ExcelReader(final String path, final String sheetName) {
        if (path == null && sheetName == null) {
            return;
        }

        this.setPath(path);

        if (this.getPath() != null) {
            try {
                setFis(new FileInputStream(path));

                if (this.getFis() != null) {
                    XSSFWorkbook xssfWB = new XSSFWorkbook(getFis());

                    if (xssfWB != null) {
                        setWorkbook(xssfWB);

                        if (this.getWorkbook() != null) {
                            this.setSheet(getWorkbook().getSheet(sheetName));
                            if (this.getSheet() == null) {
                                LOG.warn("Unable to retrieve worksheet \"" + sheetName + "\" from workbook \"" + path + "\"");
                            }
                        }
                    }
                }
            } catch (FileNotFoundException fe) {
                LOG.warn("Unable to open spreadsheet " + fe.getMessage());
            } catch (IOException ie) {
                LOG.warn("Unable to open spreadsheet " + ie.getMessage());
            } catch (Exception e) {
                LOG.error("Problem creating stream for path \"" + path + "\"", e);
            } finally {
                if (getFis() != null) {
                    try {
                        getFis().close();
                    } catch (IOException e) {
                        LOG.error(e.getStackTrace());
                    }
                }
            }

        } else {
            LOG.warn("Unable to open spreadsheet as path is null");
        }
    }

    /**
     * This method returns the String value of the cell content which matches
     * for the specified row number and column number.
     *
     * @author Sujoy
     * @param row
     *            - Row Number
     * @param col
     *            - column Number
     * @return String Note: Row number and column numbers start from index 0.
     **/
    public String getCellValue(final int row, final int col) {
        Row rows = getSheet().getRow(row - 1);
        if (rows == null){
            LOG.warn("Attempt to read data from cell in undefined row " 
                    + "[" + row + "," + col + "] on sheet \""
                    + getSheet().getSheetName() + "\"");
            return "";
        } else {
            Cell cell;
            try {
                cell = rows.getCell(col);
            } catch (IllegalArgumentException ex) {
                LOG.warn("Attempt to read data from undefined cell "
                        + "[" + row + "," + col + "] on sheet \""
                        + getSheet().getSheetName() + "\"");
                throw ex;
            }
            if (cell == null) {
                // Treat cells that don't exist as empty because Excel generally only creates the
                // cells that have values or styles
                return "";
            } else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                return cell.getStringCellValue();
            } else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                // return String.valueOf(cell.getNumericCellValue());
                return new java.text.DecimalFormat("0").format(cell.getNumericCellValue());
            } else {
                return cell.getStringCellValue();
            }
        }
    }

    /**
     * This method returns the String value of the cell content. If Date cell matches then using date format as yyyy-MM-dd.
     *
     * @param row
     *            - Row Number
     * @param col
     *            - column Number
     * @return String Note: Row number and column numbers start from index 0.
     **/
    public String getFormatedValue(final int row, final int col) {
        Row rows = getSheet().getRow(row - 1);
        Cell cell = rows.getCell(col);
        if (cell == null) {
            return "";
        }
        int type = cell.getCellType();
        Object result;
        switch (type) {

            case 0: // numeric value in Excel
                if (DateUtil.isCellInternalDateFormatted(cell)){
                    SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
                    result = dateFormat.format(cell.getDateCellValue());
                } else {
                    result = new java.text.DecimalFormat("0").format(cell
                            .getNumericCellValue());
                }
                break;
            case 1: // String Value in Excel 
                result = cell.getStringCellValue();
                break;
            default:
                result =  cell.getStringCellValue();
        }
        return result.toString();
    }

    /**
     * This overloaded method returns the String value of the cell content which
     * matches for the specified SheetName, column number and row number.
     *
     * @author Sujoy
     * @param sheetName - String value of the workSheetName
     * @param row - Row Number
     * @param col - column Number
     * @return String Note: column numbers start from index 0.
     **/
    public String getCellValue(final String sheetName, final int row, final int col) {
        setSheet(getWorkbook().getSheet(sheetName));
        if (getSheet() != null){
            Row rows = getSheet().getRow(row - 1);
            Cell cell = rows.getCell(col);

            if (cell == null) {
                return "";
            } else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                return cell.getStringCellValue();
            } else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                // return String.valueOf(cell.getNumericCellValue());
                return new java.text.DecimalFormat("0").format(cell
                        .getNumericCellValue());
            } else {
                return cell.getStringCellValue();
            }
        } else {
            LOG.warn("Unable to retrieve sheet \"" + sheetName + "\" from workbook \"" + getPath() + "\"");
            return null;
        }
    }

    /**
     * This overloaded method returns the String value of the cell content which matches for the specified SheetName, row number and column number.
     *
     * This will gets a String, Numeric or Boolean from the cell as an Object return type.
     *
     * @param sheetName - String value of the workSheetName
     * @param row - Row Number
     * @param col - column Number
     * @return Object (as a String, Double or Boolean)
     **/
    public Object getCellValueByFormat(final String sheetName, final int row, final int col) {
        if (sheetName != null) {
            if (getWorkbook() != null) {
                XSSFSheet eSheet = getWorkbook().getSheet(sheetName);
                if (eSheet != null) {
                    setSheet(eSheet);
                    if (getSheet() != null) {
                        Row rows = getSheet().getRow(row - 1);
                        Cell cell = rows.getCell(col);
                        if (cell == null) {
                            return "";
                        } else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                            return cell.getStringCellValue();
                        } else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                            return (Double) cell.getNumericCellValue();
                        } else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
                            return Boolean.valueOf(cell.getBooleanCellValue());
                        } else {
                            return cell.getStringCellValue();
                        }
                    }
                } else {
                    LOG.warn("Unable to retrieve sheet \"" + sheetName + "\" from workbook \"" + getPath() + "\"");
                    return null;
                }
            }
        }
        return null;
    }

    /**
     * This method returns the number of usedRows present in a worksheet.
     *
     * @author Sujoy
     * @return integer
     **/

    public int getRowCount() {
        if (getSheet() != null){
            int rowCount = getSheet().getLastRowNum();
            return rowCount + 1;
        } else {
            LOG.warn("Sheet not found");
            return 0;
        }
    }

    /**
     * This is a overloaded method which returns the number of usedRows present
     * in the specified workSheet.
     *
     * @author Sujoy
     * @param sheetName
     *            String value of the workSheetName
     * @return integer
     *
     **/
    public int getRowCount(final String sheetName) {
        setSheet(getWorkbook().getSheet(sheetName));
        if (getSheet() != null){
            int rowCount = getSheet().getLastRowNum() + 1;
            return rowCount;
        } else {
            LOG.warn("Sheet " + sheetName + " not found");
            return 0;
        }
    }

    /**
     * This method returns the number of columns present in a particular
     * worksheet for the specified rownumber.
     *
     * @author Sujoy
     * @return integer
     *
     **/
    public int getColumnCount() {
        if (getSheet() != null){
            Row rowNum = getSheet().getRow(0);
            int columnCount = rowNum.getLastCellNum();
            return columnCount;
        } else {
            LOG.warn("Sheet not found");
            return 0;
        }
    }

    /**
     * This overloaded method returns the number of columns present in a
     * particular worksheet for the specified worksheet.
     *
     * @author Sujoy
     * @return integer
     * @param sheetName
     *            String value of the workSheetName
     *
     **/

    // returns number of columns in a sheet
    public int getColumnCount(final String sheetName) {
        setSheet(getWorkbook().getSheet(sheetName));
        if (getSheet() != null){
            Row rowNum = getSheet().getRow(0);
            int columnCount = rowNum.getLastCellNum();
            return columnCount;
        } else {
            LOG.warn("Sheet " + sheetName + " not found");
            return 0;
        }
    }

    /**
     * This method is to set String value in Excel file.
     * 
     * @param sheetName is workSheet Name in which value need to be set
     * @param colName is to find particular column to set value
     * @param rowNum is to find particular cell to set value
     * @param data is the string to set
     * @return Boolean whether data added or not
     */
    public boolean setCellData(final String sheetName, final String colName, final int rowNum, final String data) {
        try {
            if (getPath() != null) {
                setFis(new FileInputStream(getPath()));

                if (getFis() != null) {
                    setWorkbook(new XSSFWorkbook(getFis()));
                    if (getWorkbook() != null) {
                        if (rowNum <= 0) {
                            return false;
                        }

                        int index = getWorkbook().getSheetIndex(sheetName);
                        int colNum = -1;
                        if (index == -1) {
                            return false;
                        }

                        setSheet(getWorkbook().getSheetAt(index));

                        setRow(getSheet().getRow(0));
                        for (int i = 0; i < getRow().getLastCellNum(); i++) {
                            if (getRow().getCell(i).getStringCellValue().trim().equals(colName)) {
                                colNum = i;
                            }
                        }

                        if (colNum == -1) {
                            return false;
                        }

                        getSheet().autoSizeColumn(colNum);
                        setRow(getSheet().getRow(rowNum - 1));

                        if (getRow() == null) {
                            setRow(getSheet().createRow(rowNum - 1));
                        }

                        setCell(getRow().getCell(colNum));

                        if (getCell() == null) {
                            setCell(getRow().createCell(colNum));
                        }

                        getCell().setCellValue(data);

                        setFileOut(new FileOutputStream(getPath()));

                        getWorkbook().write(getFileOut());

                    }
                }
            }
        } catch (Exception e) {
            LOG.error("Failed to set cell value", e);
            return false;
        } finally {
            if (getFis() != null) {
                try {
                    getFis().close();
                } catch (IOException e) {
                    LOG.error(e.getStackTrace());
                }
            }

            if (getFileOut() != null) {
                try {
                    getFileOut().close();
                } catch (IOException e) {
                    LOG.error(e.getStackTrace());
                }
            }
        }

        return true;
    }

    /**
     * Getter.
     * @return the path
     */
    public String getPath() {
        return path;
    }

    /**
     * Setter.
     * @param path the path to set
     */
    public void setPath(final String path) {
        this.path = path;
    }

    /**
     * Getter.
     * @return the fis
     */
    public FileInputStream getFis() {
        return fis;
    }

    /**
     * Setter.
     * @param fis the fis to set
     */
    public void setFis(final FileInputStream fis) {
        this.fis = fis;
    }

    /**
     * Getter.
     * @return the workbook
     */
    public XSSFWorkbook getWorkbook() {
        return workbook;
    }

    /**
     * Setter.
     * @param workbook the workbook to set
     */
    public void setWorkbook(final XSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    /**
     * Getter.
     * @return the sheet
     */
    public XSSFSheet getSheet() {
        return sheet;
    }

    /**
     * Setter.
     * @param sheet the sheet to set
     */
    public void setSheet(final XSSFSheet sheet) {
        this.sheet = sheet;
    }

    /**
     * Getter.
     * @return the row
     */
    public XSSFRow getRow() {
        return row;
    }

    /**
     * Setter.
     * @param row the row to set
     */
    public void setRow(final XSSFRow row) {
        this.row = row;
    }

    /**
     * Getter.
     * @return the cell
     */
    public XSSFCell getCell() {
        return cell;
    }

    /**
     * Setter.
     * @param cell the cell to set
     */
    public void setCell(final XSSFCell cell) {
        this.cell = cell;
    }

    /**
     * Getter.
     * @return the fileOut
     */
    public FileOutputStream getFileOut() {
        return fileOut;
    }

    /**
     * Setter.
     * @param fileOut the fileOut to set
     */
    public void setFileOut(final FileOutputStream fileOut) {
        this.fileOut = fileOut;
    }

    
    
    /**
     * getRowData - reads the data from spreadsheet for the matching rowNumber specified
     * and returns the data as a Map collection with the rowHeaders of the work-sheet as key
     * and the corresponding row as value.
     *
     * @author Sujoy
     * @param rownumber - rowN Number for which the values needs to be mapped to Column Headers
     * @return {@link Map}
     */
    public Map<String, String> getRowData(final int rownumber){
        String columnheader;
        String columnvalue;
        int colCount = this.getColumnCount();

        HashMap<String, String> map = new HashMap<String, String>();

        for (int colNum = 0; colNum < colCount; colNum++) {
            columnheader = this.getCellValue(1, colNum);
            columnvalue = this.getCellValue(rownumber, colNum);
            map.put(columnheader, columnvalue);
        }

        return map;
    }
    
    
	
	/**
	 * 
	 * Reads the data from spreadsheet and returns the data as a dynamic Object
	 * 
	 * array.
	 * @author Sujoy
	 * @param excelWorkBookName
	 * 
	 *                          - excel sheets name placed in the "Data" folder
	 * 
	 * @param workSheetName
	 * 
	 *                          - work sheets name
	 * 
	 * @return Object[][]
	 * 
	 */

    public  Object[][] getDataFromSpreadSheet() {

        
        int rowCount = getRowCount();

        int colCount = getColumnCount();

        Object[][] data = new Object[rowCount - 1][colCount];

        for (int rowNum = 2; rowNum <= rowCount; rowNum++) {

            // loop all the available row values

            for (int colNum = 0; colNum < colCount; colNum++) {

                // while returning the data[][] you should not send the header

                // values

                data[rowNum - 2][colNum] = getCellValue(rowNum, colNum);

            }

        }

 

        return data;

    }
}
