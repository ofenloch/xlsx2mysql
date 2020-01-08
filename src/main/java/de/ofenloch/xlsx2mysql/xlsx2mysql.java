package de.ofenloch.xlsx2mysql;

import java.beans.Statement;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.apache.commons.io.FilenameUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class xlsx2mysql {

  protected static final Logger log = LogManager.getLogger(xlsx2mysql.class.getName());
  /**
   * file name of the excel file
   */
  String xlsxFileName;
  /**
   * the Wokrbook
   */
  Workbook workBook;
  /**
   * index of the row with the column names (defaults to 0)
   */
  List<Integer> iColumnNamesRowIdx = null;
  /**
   * index of the first data row (defaults to 1)
   */
  List<Integer> iFirstDataRowIdx = null;
  /**
   * the name of the database to be created (defaults to the file name of the
   * Excel file) (the db will be deleted and re-created if it exists)
   */
  String dbName;
  /**
   * the name of the table to be created (defaults to the name of the first sheet)
   */
  List<String> tblName = null;
  /**
   * the names of the columns
   */
  // List<Column> columnNames;
  HashMap<Integer, List<Column>> columnNames;
  /**
   * the number of sheets in the workbook
   */
  int nSheets;
  /**
   * replace newline characters in Excel names with this:
   */
  static private String newLineReplacement = "<newline>";

  /**
   * class to handel columns in Excel and MySQL
   */
  public static class Column {
    /**
     * column's index
     */
    private int columnIndex;
    /**
     * the sheet's index this column belongs to
     */
    private int sheetIndex;
    /**
     * column's name in the database
     */
    private String sqlName;
    /**
     * column's original in the Excel doument
     */
    private String excelName;
    /**
     * cloumn's original index in the Excel document
     */
    private int excelIndex;
    /**
     * column's type in the Excel document
     */
    private CellType excelType;

    /**
     * default cunstructor
     *
     * @param cell
     * @param columnIndex
     * @param sheetIndex
     */
    public Column(Cell cell, int columnIndex, int sheetIndex) {
      this.columnIndex = columnIndex;
      this.sheetIndex = sheetIndex;
      this.excelName = getCellValueAsString(cell);
      this.sqlName = sanitizeColumnName(excelName);
      this.excelType = cell.getCellType();
      this.excelIndex = cell.getColumnIndex();
    }

    /**
     * construct a new Column object and set its sqlName attribute to the given
     * value
     *
     * @param cell
     * @param columnIndex
     * @param sheetIndex
     * @param sqlName
     */
    public Column(Cell cell, int columnIndex, int sheetIndex, String sqlName) {
      this(cell, columnIndex, sheetIndex);
      this.sqlName = sanitizeColumnName(sqlName);
    }

    public String getSqlName() {
      return this.sqlName;
    }

    public String getExcelName() {
      return this.excelName;
    }

    public int getExcelIndex() {
      return this.excelIndex;
    }

    public CellType getExcelType() {
      return this.excelType;
    }

    /**
     *
     * @return textual value for the cell type
     *
     *         Values are taken from class CellType and may need update after
     *         changing the POI version.
     */
    public String getExcelTypeName() {
      String typeName;
      switch (this.excelType) {
      case NUMERIC:
        /**
         * Numeric cell type (whole numbers, fractional numbers, dates)
         */
        typeName = "NUMERIC(0)";
        break;
      case STRING:
        /** String (text) cell type */
        typeName = "STRING(1)";
        break;
      case FORMULA:
        /**
         * Formula cell type
         *
         * @see FormulaType
         */
        typeName = "FORMULA(2)";
        break;
      case BLANK:
        /**
         * Blank cell type
         */
        typeName = "BLANK(3)";
        break;
      case BOOLEAN:
        /**
         * Boolean cell type
         */
        typeName = "BOOLEAN(4)";
        break;
      case ERROR:
        /**
         * Error cell type
         *
         * @see FormulaError
         */
        typeName = "ERROR(5)";
        break;
      default:
        /**
         * Unknown type, used to represent a state prior to initialization or the lack
         * of a concrete type. For internal use only.
         */
        // @Internal(since="POI 3.15 beta 3")
        typeName = "NONE(-1)";
        break;
      } // switch (this.excelType)
      return typeName;
    }

    public static String sanitizeColumnName(String originalName) {
      String sanitzedName = originalName.replace(" ", "_");
      sanitzedName = sanitzedName.replace("\n", "");
      sanitzedName = sanitzedName.replace(newLineReplacement, "");
      sanitzedName = sanitzedName.replace("`", "``");
      sanitzedName = sanitzedName.replace("\"", "``");
      sanitzedName = sanitzedName.replace("\\", "\\\\");

      return sanitzedName.trim();
    }

    /**
     * returns the cell's value as string no matter what type it realy is
     *
     * @param cell
     * @return
     */
    public static String getCellValueAsString(Cell cell) {
      String value = "";
      CellType cellType = cell.getCellType();
      if (cellType == CellType.NUMERIC) {
        value = Double.toString(cell.getNumericCellValue());
      } else if (cellType == CellType.BOOLEAN) {
        value = cell.getBooleanCellValue() == true ? "true" : "false";
      } else if (cellType == CellType.FORMULA) {
        value = cell.getCellFormula();
        // TODO: should we do an evaluation of formulae?
        // See https://poi.apache.org/components/spreadsheet/eval.html
      } else {
        value = cell.getStringCellValue();
      }
      // It might be better to replace at least the newline
      // characters in the Excel name:
      value = value.replaceAll("\n", newLineReplacement);
      // try something like this
      // value = value.replaceAll("\n", "\\\\\n");
      return value;
    } // public static String getCellValueAsString(Cell cell)

  } // public static class Column

  /**
   * String constant for missing cells
   */
  static final public String ExcelMissingData = "N/A";
  /**
   * String constant for error cells
   */
  static final public String ExcelErrorData = "ERROR IN XLS(X)";

  /**
   * default c'tor gets only a file name
   *
   * @param fileName
   */
  public xlsx2mysql(String fileName) {
    this(fileName, FilenameUtils.getBaseName(fileName));
  } // public xlsx2mysql(String fileName)

  /**
   *
   * @param fileName
   * @param databaseName
   */
  public xlsx2mysql(String fileName, String databaseName) {
    try {
      this.xlsxFileName = fileName;
      this.dbName = databaseName;
      log.info(" xls(x) file is \"" + this.xlsxFileName + "\"");
      log.info(" database name is \"" + this.dbName + "\"");
      try {
        File f = new File(fileName);
        FileInputStream fis = new FileInputStream(f);
        if (fileName.endsWith(".xls")) {
          workBook = new HSSFWorkbook(fis);
        } else {
          workBook = new XSSFWorkbook(fis);
        }
        this.nSheets = workBook.getNumberOfSheets();
        this.tblName = new ArrayList<String>(nSheets);
        this.iColumnNamesRowIdx = new ArrayList<Integer>(nSheets);
        this.iFirstDataRowIdx = new ArrayList<Integer>(nSheets);
        this.columnNames = new HashMap<Integer, List<Column>>(nSheets);

        // initialize the lists with their defaults
        for (int i = 0; i < nSheets; i++) {
          this.iColumnNamesRowIdx.add(i, 0);
          this.iFirstDataRowIdx.add(i, 1);
          String sheetName = workBook.getSheetName(i);
          log.debug("sheet " + i + " is named \"" + sheetName + "\"");
          this.tblName.add(i, Column.sanitizeColumnName(sheetName));
          log.debug("table " + i + " is named \"" + this.tblName.get(i) + "\"");
        }
      } catch (Exception ex) {
        exceptionHandler("xlsx2mysql c'tor", ex);
      }

    } catch (Exception ex) {
      exceptionHandler("xlsx2mysql c'tor", ex);
    }
  }

  /**
   * set the index of the first row with data
   *
   * @param zeroBasedRowIndex
   */
  public void setFirstDataRowIndex(int sheetIndex, int zeroBasedRowIndex) {
    iFirstDataRowIdx.add(sheetIndex, zeroBasedRowIndex);
  }

  /**
   * get the index of the first row with data
   */
  public int getFirstDataRowIndex(int sheetIndex) {
    return iFirstDataRowIdx.get(sheetIndex);
  }

  /**
   * set the index of the row that contains the colum names
   *
   * @param zeroBasedRowIndex
   */
  public void setColumnNamesRowIndex(int sheetIndex, int zeroBasedRowIndex) {
    iColumnNamesRowIdx.add(sheetIndex, zeroBasedRowIndex);
  }

  /**
   * get the index of the row that contains the colum names
   */
  public int getColumnNamesRowIndex(int sheetIndex) {
    return iColumnNamesRowIdx.get(sheetIndex);
  }

  /**
   * generate the sql script to create the databse, the tables and fill them with
   * data in the xls(x) document
   *
   * @return
   */
  public String generateSqlScript() {
    for (int i = 0; i < nSheets; i++) {
      readColumnNames(i);
    }
    StringBuilder theScript = new StringBuilder();
    theScript.append(generateScriptHeader());
    theScript.append(generateCreateDBStatement());
    for (int i = 0; i < nSheets; i++) {
      theScript.append(generateSqlCreateTableStatement(i));
    }
    for (int i = 0; i < nSheets; i++) {
      List<String> insertStatements = new ArrayList<String>();
      theScript.append(generateSqlInsertStatements(i));
    }
    return theScript.toString();
  }

  /**
   * 
   * @param originalName
   * @return
   */
  public static String sanitizeCellValues(String originalName) {
    String sanitzedName = originalName.replace("\n", "");
    sanitzedName = sanitzedName.replace(newLineReplacement, "");
    sanitzedName = sanitzedName.replace("`", "``");
    sanitzedName = sanitzedName.replace("\"", "``");
    sanitzedName = sanitzedName.replace("\\", "\\\\");
    return sanitzedName.trim();
  }

  /**
   * get the column names in the given sheet
   */
  private void readColumnNames(int sheetIdx) {
    log.debug("readColumnNames: reading column names of sheet " + sheetIdx + " ...");
    ArrayList<Column> columnList = new ArrayList<Column>();
    int iCurrentSheet = 0;
    int iCurrentRow = 0;
    int iCurrentCell = 0;
    Set<String> names = new HashSet<>();
    for (Sheet sheet : workBook) {
      iCurrentRow = 0;
      for (Row row : sheet) {
        iCurrentCell = 0;
        if (iCurrentSheet == sheetIdx && iCurrentRow == iColumnNamesRowIdx.get(sheetIdx)) {
          for (Cell cell : row) {
            String name = Column.getCellValueAsString(cell);
            if (!names.add(name)) {
              // the name is already in the list, so we change it
              name += "_new";
              columnList.add(new Column(cell, iCurrentCell, iCurrentSheet, name));
            } else {
              columnList.add(new Column(cell, iCurrentCell, iCurrentSheet));
            }
            iCurrentCell++;
          } // for (Cell cell : row)
          columnNames.put(sheetIdx, columnList);
          log.debug("readColumnNames: read " + columnList.size() + " column names of sheet " + sheetIdx + ".");
          return;
        } // if (iCurrentSheet == sheetIdx && iCurrentRow == iColumnNamesRowIdx)
        iCurrentRow++;
      } // for (Row row : sheet)
      iCurrentSheet++;
    } // for (Sheet sheet : workBook)
  } // private void readColumnNames()

  /**
   * generate the first lines of the sql script
   *
   * @return the first lines of the script as a String
   */
  private String generateScriptHeader() {
    log.debug("generateScriptHeader: generating script header ...");
    StringBuilder theHeader = new StringBuilder();
    theHeader.append("-- ***************************************************************************************\n");
    theHeader.append("-- ---------------------------------------------------------------------------------------\n");
    theHeader.append("-- WARNING: THERE IS ABSOLUTELY NO WARRANTY THAT THIS SCRIPT DOESN'T KILL YOUR DATABASE!!!\n");
    theHeader.append("-- ---------------------------------------------------------------------------------------\n");
    theHeader.append("-- ***************************************************************************************\n");
    theHeader.append("--\n");
    theHeader.append("-- MySQL script to import data in Excel document \n");
    theHeader.append("--               \"" + xlsxFileName + "\"\n");
    theHeader.append("-- into database \n");
    theHeader.append("--               \"" + dbName + "\"\n");
    theHeader.append("--\n");
    theHeader.append("-- ***************************************************************************************\n");
    theHeader.append("--\n");
    theHeader.append("-- Depending on your data and its qualitiy you should revise tis script befor running it!\n");
    theHeader.append("--\n");
    theHeader.append("-- Once You are done with editing this script You can run it by calling something like\n");
    theHeader.append("--      mysql `cat ~/.mysql/user@dbhost` < thisFileName.sql \n");
    theHeader.append("--\n");
    theHeader.append("-- ***************************************************************************************\n");
    theHeader.append("--\n");
    theHeader.append("-- Check the script carefully before running it!\n");
    theHeader.append("--\n");
    theHeader.append("--   * There is only a very simple check for duplicate column names.\n");
    theHeader.append("--   * There is no check for invalid column names.\n");
    theHeader.append("--   * There is no check for complete INSERT statements.\n");
    theHeader.append("--   * There is absolutely no warranty that the script doesn't kill your database!\n");
    theHeader.append("--\n");
    if (nSheets == 1) {
      theHeader.append("-- The document contains only one sheet.\n");
    } else {
      theHeader.append("-- The document contains " + nSheets + " sheets:\n");
      for (int i = 0; i < nSheets; i++) {
        theHeader.append("--     sheet \"" + workBook.getSheetName(i) + "\" -> table \"" + tblName.get(i) + "\"\n");
      }
    }
    theHeader.append("--\n");
    log.debug("generateScriptHeader: done generating script header.");
    return theHeader.toString();
  } // private String generateScriptHeader()

  /**
   * generate the sql script snippet to create the database
   *
   * @return the script snippet as a String
   */
  private String generateCreateDBStatement() {
    log.debug("generateCreateDBStatement: generating CREATE DATABASE statement ...");
    StringBuilder stringBuilder = new StringBuilder();
    stringBuilder.append("-- Delete the database if it exists\n");
    stringBuilder.append("DROP DATABASE IF EXISTS `" + dbName + "`; " + "\n");
    stringBuilder.append("-- (Re-) Create the database\n");
    stringBuilder.append("CREATE DATABASE IF NOT EXISTS `" + dbName + "` " + "\n");
    stringBuilder.append(" DEFAULT CHARACTER SET utf8" + "\n");
    stringBuilder.append(" DEFAULT COLLATE utf8_general_ci;" + "\n");
    stringBuilder.append("-- Switch to the new database\n");
    stringBuilder.append("USE `" + dbName + "`;" + "\n");
    log.debug("generateCreateDBStatement: done generating CREATE DATABASE statement.");
    return stringBuilder.toString();
  } // private String generateCreateDBStatement()

  /**
   * generate the sql script snippet for generating the table corresponding to the
   * given sheet
   *
   * @param sheetIndex the index of the selected sheet
   * @return the script snippet as a String
   */
  private String generateSqlCreateTableStatement(int sheetIndex) {
    log.debug("generateSqlCreateTableStatement: generating CREATE TABLE statement for table " + sheetIndex + " ...");
    final int nColumnNames = columnNames.get(sheetIndex).size();
    final int columnNamesRowIdx = iColumnNamesRowIdx.get(sheetIndex);
    final int firstDataRowIdx = iFirstDataRowIdx.get(sheetIndex);
    StringBuilder stringBuilder = new StringBuilder();
    stringBuilder.append("-- \n");
    stringBuilder.append("-- create table for sheet " + sheetIndex + "\n");
    stringBuilder.append("--   found " + nColumnNames + " column names" + "\n");
    stringBuilder.append("--   column names taken from row with idx " + columnNamesRowIdx + "\n");
    stringBuilder.append("--   data start in row with idx " + firstDataRowIdx + "\n");
    stringBuilder.append("-- \n");
    stringBuilder.append("DROP TABLE IF EXISTS `" + dbName + "`.`" + tblName.get(sheetIndex) + "`;" + "\n");
    stringBuilder.append("CREATE TABLE `" + dbName + "`.`" + tblName.get(sheetIndex) + "` (" + "\n");
    for (Column column : columnNames.get(sheetIndex)) {
      stringBuilder.append(" `" + column.getSqlName() + "` VARCHAR(200) NOT NULL DEFAULT '" + ExcelMissingData + "' ");
      // The comment with the Excel name of the coumn may cause trouble if it contains
      // some "weird" characters. So we put it only as comment into the script.
      stringBuilder.append("/* COMMENT ` ExcelName \'" + column.getExcelName() + "\' , ExcelType ");
      stringBuilder.append(column.getExcelTypeName() + " ExcelColumnIndex " + column.getExcelIndex() + " ` */");
      stringBuilder.append(", \n");
    }
    String statement = stringBuilder.toString();
    // remove the \n and the last comma to avoid MySQL errors
    statement = statement.substring(0, statement.length() - 3);
    statement += "\n);\n";
    log.debug(
        "generateSqlCreateTableStatement: done generating CREATE TABLE statement for table " + sheetIndex + " ...");
    return statement;
  } // private String generateSqlCreateTableStatement(int sheetIndex)

  /**
   * generate the sql script for inserting data from the selected sheet into the
   * right table
   *
   * @param sheetIndex the index of the selected sheet
   * @return the script as a String
   */
  private String generateSqlInsertStatements(int sheetIndex) {
    final int nColumnNames = columnNames.get(sheetIndex).size();
    final int columnNamesRowIdx = iColumnNamesRowIdx.get(sheetIndex);
    final int firstDataRowIdx = iFirstDataRowIdx.get(sheetIndex);
    log.debug("generateSqlInsertStatements: generating INSERT statements for table " + sheetIndex + " ...");
    StringBuilder stringBuilder = new StringBuilder();
    stringBuilder.append("-- \n");
    stringBuilder.append("-- insert data of sheet " + sheetIndex + "\n");
    stringBuilder.append("--   found " + nColumnNames + " column names" + "\n");
    stringBuilder.append("--   column names taken from row with idx " + columnNamesRowIdx + "\n");
    stringBuilder.append("--   data start in row with idx " + firstDataRowIdx + "\n");
    stringBuilder.append("-- \n");

    int iCurrentSheet = 0;
    int iCurrentRow = 0;
    int iCurrentCell = 0;

    for (Sheet sheet : workBook) {
      if (iCurrentSheet == sheetIndex) {
        iCurrentRow = 0;

        for (Row row : sheet) {
          iCurrentCell = 0;
          if (iCurrentRow >= firstDataRowIdx) {
            stringBuilder.append("INSERT INTO `" + dbName + "`.`" + tblName.get(iCurrentSheet) + "` VALUES (" + "\n ");
            for (iCurrentCell = 0; iCurrentCell < nColumnNames; iCurrentCell++) {
              // In cases like this
              //
              // __A__B__C__D
              // 1 A1 B1 C1 D1
              // 2 A2 __ C2 D2
              // 3 __ B3 C3 D3
              //
              // Excel does not create all cells. So cells B2 and A3 would be missing. To get
              // the right cell into the right column of our DB, we have to loop over cells in
              // a row from 0 to nColumnNames and check if row.getCell(i) returns a cell or
              // null. In this exymple, we'd get null for cells B2 and A3 and need to handle
              // it properly.
              //
              // After fixing this, ther should be no need to fill missing cells. So we throw
              // an exception if this happens anyways.
              //
              // Also, there should be no need to drop the NOT NULL in the CREATE TABLE
              // statment.
              Cell cell = row.getCell(iCurrentCell);
              if (cell != null) {
                CellType cellType = cell.getCellType();
                if (cellType == CellType.NUMERIC) {
                  double number = cell.getNumericCellValue();
                  stringBuilder.append(number + ", ");
                } else if (cellType == CellType.ERROR) {
                  stringBuilder.append("\"" + ExcelErrorData + "\", ");
                } else {
                  String cellVal = Column.getCellValueAsString(cell);
                  if (cellVal.length() < 1) {
                    cellVal = ExcelMissingData;
                  }
                  stringBuilder.append("\"" + sanitizeCellValues(cellVal) + "\", ");
                }
              } else {
                // There is no cell event though there should be one!
                stringBuilder.append("\"" + ExcelMissingData + "\", ");
              }
            } // for (iCurrentCell=0; iCurrentCell<nColumnNames; iCurrentCell++ )
            // fill up missing cells
            for (int i = iCurrentCell; i < nColumnNames; i++) {
              // Depending on the cell type we'd have to append a
              // a String or a Number or something else.
              // So, we append a null value instead of examining the type.
              // Thus, we skip the NOT NULL in the CREATE TABLE statment.
              // stringBuilder.append("\"" + ExcelMissingData + "\", ");
              stringBuilder.append("null, ");
              throw new RuntimeException("generateSqlInsertStatements: This should never happen!");
            }
            // remove the last ", " to avoid MySQL errors
            stringBuilder.deleteCharAt(stringBuilder.lastIndexOf(","));
            stringBuilder.append(");\n");
          } // if (iCurrentRow >= firstDataRowIdx)
          // log.debug("generateSqlInsertStatements: row #" + iCurrentRow + " done.");
          iCurrentRow++;
        } // for (Row row : sheet)
        break;
      } // if (iCurrentSheet == sheetIndex)
      iCurrentSheet++;
    } // for (Sheet sheet : workBook)
    log.debug("generateSqlInsertStatements: done generating INSERT statements for table " + sheetIndex + " ...");
    return stringBuilder.toString();
  } // private String generateSqlInsertStatements()

  /**
   * Exception handler that simply logs a message
   */
  public static void exceptionHandler(String origin, Exception ex) {
    try {
      exceptionHandler(origin, ex, false);
    } catch (Exception e) {
      // Nothing to do here. The try/catch is only neede to satisfy the compiler
      // since the original exception handler had to get a throws declaration.
    }
  } // public static void exceptionHandler(String origin, Exception ex)

  /**
   * Exception handler that simply logs a message
   *
   * Rethrows exception if rethrow is true
   *
   * @throws Exception
   */
  public static void exceptionHandler(String origin, Exception ex, boolean rethrow) throws Exception {
    StringWriter message = new StringWriter();
    message.append("caught exception in " + origin + " : ");
    message.append("  message: " + ex.getMessage());
    message.append("  cause  : " + ex.getCause());
    message.append("  class  : " + ex.getClass().getName());
    message.append("  statck trace  : ");
    StringWriter stackTrace = new StringWriter();
    ex.printStackTrace(new PrintWriter(stackTrace));
    message.append(stackTrace.toString());
    String sMessage = message.toString();
    log.error(sMessage);
    System.err.println(sMessage);
    if (rethrow == true) {
      throw ex;
    }
  } // public static void exceptionHandler(String origin, Exception ex)

  /**
   * calculate for an Excel column name the column number
   * 
   * e.g. A -> 1, Z -> 26, AA -> 27, ..., AAA -> 703, ....
   * 
   * @param s the column's Excel name (e.g. "AT")
   * @return the column's number (1-based, not the index)
   */
  static public int ColumnName2Number(String s) {
    int result = 0;
    for (char c : s.toCharArray()) {
      result = result * 26 + (c - 'A') + 1;
    }
    return result;
  }

  /**
   * generate an Excel colum name from the given column number
   * 
   * e.g. 1 -> A, 25 -> Y, 27 -> AA, ..., 703 -> AAA, ...
   * 
   * @param n the column number
   * @return the Excel column name to the given number
   */
  static public String ColumnNumberToName(int n) {
    if (n <= 0) {
      throw new IllegalArgumentException("Input is not valid!");
    }

    StringBuilder sb = new StringBuilder();

    while (n > 0) {
      n--;
      char ch = (char) (n % 26 + 'A');
      n /= 26;
      sb.append(ch);
    }

    sb.reverse();
    return sb.toString();
  }

} // public class xlsx2mysql