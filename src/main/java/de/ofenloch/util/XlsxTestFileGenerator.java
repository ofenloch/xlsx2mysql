package de.ofenloch.util;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Random;

import javax.management.RuntimeErrorException;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import de.ofenloch.util.RandomString;

public class XlsxTestFileGenerator {

    public static Random rnd = new Random();

    public static final Logger log = LogManager.getLogger(XlsxTestFileGenerator.class.getName());

    public static void main(String[] args) {
        String fileName = "test.xlsx";
        int nSheets = 5;
        int nMaxRows = 100;
        int nMaxColumns = 50;
        // get command line parameters
        final int nArgs = args.length;
        if (nArgs > 0) {
            fileName = args[0];
            if (nArgs > 1) {
                nSheets = Integer.parseInt(args[1]);
                if (nArgs > 2) {
                    nMaxRows = Integer.parseInt(args[2]);
                    if (nArgs > 3) {
                        nMaxColumns = Integer.parseInt(args[3]);
                    }
                }
            }
        }
        System.out.println("creating file \"" + fileName + "\" with " + nSheets + " sheets");
        System.out.println("     nMaxRows : " + nMaxRows);
        System.out.println("  nMaxColumns : " + nMaxColumns);
        generateTestFile(fileName, nSheets, nMaxRows, nMaxColumns);
        System.out.println("done!");
    }

    public static int generateTestFile(final String fileName, final int nSheets, final int nMaxRows,
            final int nMaxColums) {
        int error = 0;

        try {
            Workbook wb = null;

            if (fileName.endsWith(".xlsx")) {
                // for the new format (xlsx files) we use XSSFWorkbook
                log.info("creating xlsx workbook \"" + fileName + "\" ...");
                wb = new XSSFWorkbook();
            } else if (fileName.endsWith(".xls")) {
                // for the old format (xls files) we use HSSFWorkbook
                log.info("creating xls workbook \"" + fileName + "\" ...");
                wb = new HSSFWorkbook();
            } else {
                log.error("don't know how to create file \"" + fileName + "\".");
                throw new IllegalArgumentException("don't know how to create file \"" + fileName + "\".");
            }
            log.debug(" nSheets " + nSheets);
            log.debug(" nMaxRows " + nMaxRows);
            log.debug(" nMaxColums " + nMaxColums);
            for (int iSheet = 0; iSheet < nSheets; iSheet++) {
                String sheetName = "Sheet " + iSheet;
                if (iSheet % 2 == 0) {
                    sheetName = RandomString.getAlphaNumericString(12);
                } else {
                    sheetName = RandomString.getNonAlphaNumericString(12);
                }
                // You can use org.apache.poi.ss.util.WorkbookUtil#createSafeSheetName(String
                // nameProposal)}
                // for a safe way to create valid names, this utility replaces invalid
                // characters with a space (' ')
                // String safeName = WorkbookUtil.createSafeSheetName("[O'Brien's sales*?]"); //
                // returns " O'Brien's sales "
                sheetName = WorkbookUtil.createSafeSheetName(sheetName);
                Sheet sheet = wb.createSheet(sheetName);
                int nRows = 1 + rnd.nextInt(nMaxRows - 1);
                int nColums = 1 + rnd.nextInt(nMaxColums - 1);
                fillSheetWithRandomData(sheet, nRows, nColums);
                wb.close();
                if (fileName.endsWith(".xlsx")) {
                    log.debug("created xlsx workbook \"" + fileName + "\" with " + nSheets + " sheets.");
                } else {
                    log.debug("created xls workbook \"" + fileName + "\" with " + nSheets + " sheets.");
                }
            } // for (int iSheet = 0; iSheet < nSheets; iSheet++)

            OutputStream fileOut = new FileOutputStream(fileName);
            wb.write(fileOut);
            wb.close();
        } catch (Exception ex) {
            error = 10;
            ex.printStackTrace();
        }
        return error;
    } // public static int generateTestFile(final String fileName, final int nSheets,
      // final int nMaxRows,final int nMaxColums)

    public static void fillSheetWithRandomData(Sheet sheet, int nRows, int nColums) {
        log.debug("filling sheet \"" + sheet.getSheetName() + "\" with " + nRows + " rows with " + nColums
                + " colums each ...");
        Workbook wb = sheet.getWorkbook();
        String sheetName = sheet.getSheetName();
        CellStyle headerStyle = wb.createCellStyle();
        CreationHelper createHelper = wb.getCreationHelper();
        Font headerFont = wb.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 12);
        headerStyle.setFont(headerFont);
        for (int iRow = 0; iRow < nRows; iRow++) {
            Row row = sheet.createRow(iRow);
            if (iRow == 0) {
                // a row with column names
                for (int iColumn = 0; iColumn < nColums; iColumn++) {
                    Cell cell = row.createCell(iColumn);
                    cell.setCellStyle(headerStyle);
                    if (iColumn % 4 == 0) {
                        cell.setCellValue(RandomString.getNonAlphaNumericString(15));
                    } else {
                        cell.setCellValue(sheetName + "_" + iColumn + 1);
                    }
                } // for (iColumn = 0; iColumn < nColums; iColumn++)
            } else {
                // a row with random data
                for (int iColumn = 0; iColumn < nColums; iColumn++) {
                    int cellType = rnd.nextInt(6);
                    Cell cell = row.createCell(iColumn);
                    switch (cellType) {
                    case 0: // "NUMERIC(0)"
                        cell.setCellValue(iColumn % 2 == 0 ? rnd.nextDouble() : rnd.nextInt());
                        break;
                    case 4: // "BOOLEAN(4)"
                        cell.setCellValue(iColumn % 2 == 0 ? true : false);
                        break;
                    default:
                        if (iColumn % 3 == 0) {
                            cell.setCellValue(createHelper
                                    .createRichTextString(sheetName + " row " + iRow + 1 + " column " + iColumn + 1));
                        } else {
                            cell.setCellValue(RandomString.getAlphaNumericString(35));
                        }
                    } // switch (cellType)

                } // for (iColumn = 0; iColumn < nColums; iColumn++)
            }

        } // for (iRow = 0; iRow < nRows; iRow++)
    } // public static void fillSheetWithRandomData(Sheet sheet, int nRows, int
      // nColums)

} // public class XlsxTestFileGenerator