package tr.com.ramazan.fatura_entegrasyon.excel.operations;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import tr.com.ramazan.fatura_entegrasyon.interfaces.IExcelOperations;

import java.io.*;
import java.util.Iterator;
import java.util.List;

/**
 * Created by ramazancesur on 13/06/2020.
 */
public class ExcelOperations implements IExcelOperations {
    public void readXLSFile(String excelPath) throws IOException {
        InputStream ExcelFileToRead = new FileInputStream(excelPath);
        HSSFWorkbook wb = new HSSFWorkbook(ExcelFileToRead);

        HSSFSheet sheet = wb.getSheetAt(0);
        HSSFRow row;
        HSSFCell cell;

        Iterator rows = sheet.rowIterator();

        while (rows.hasNext()) {
            row = (HSSFRow) rows.next();
            Iterator cells = row.cellIterator();

            while (cells.hasNext()) {
                cell = (HSSFCell) cells.next();

                if (cell.getCellType() == CellType.STRING) {
                    System.out.print(cell.getStringCellValue() + " ");
                } else if (cell.getCellType() == CellType.NUMERIC) {
                    System.out.print(cell.getNumericCellValue() + " ");
                } else {
                    System.out.print("unexpected format");
                    //U Can Handel Boolean, Formula, Errors
                }
            }
            System.out.println();
        }

    }

    public void writeXLSFile(String excelFilePath) throws IOException {

        String excelFileName = excelFilePath;//name of excel file

        String sheetName = "Sheet1";//name of sheet

        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet(sheetName);

        //iterating r number of rows
        for (int r = 0; r < 5; r++) {
            HSSFRow row = sheet.createRow(r);

            //iterating c number of columns
            for (int c = 0; c < 5; c++) {
                HSSFCell cell = row.createCell(c);

                cell.setCellValue("Cell " + r + " " + c);
            }
        }

        FileOutputStream fileOut = new FileOutputStream(excelFileName);

        //write this workbook to an Outputstream.
        wb.write(fileOut);
        fileOut.flush();
        fileOut.close();
    }

    public void updateXLSXFile(String excelFileName, List<String> satisValues, int colomnNumber) throws IOException {
        FileInputStream inputStream = new FileInputStream(new File(excelFileName));
        XSSFWorkbook wb = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = wb.getSheetAt(0);
        for (int i = 0; i < satisValues.size(); i++) {
            String satisValue = satisValues.get(i);
            XSSFRow row = sheet.getRow(i);
            XSSFCell cell = row.createCell(colomnNumber);
            cell.setCellValue(satisValue);
        }

        inputStream.close();

        FileOutputStream outputStream = new FileOutputStream(excelFileName.replace(".xlsx", "_satisDuzenli"
                + System.currentTimeMillis() + ".xlsx"), true);
        wb.write(outputStream);
        wb.close();
        outputStream.close();
    }


    public void writeXLSXFile(String excelFilePath) throws IOException {

        String excelFileName = excelFilePath; //name of excel file

        String sheetName = "Sheet1";//name of sheet

        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(sheetName);

        sheet.getTables().get(1);
        //iterating r number of rows
        for (int r = 0; r < 5; r++) {
            XSSFRow row = sheet.createRow(r);

            //iterating c number of columns
            for (int c = 0; c < 5; c++) {
                XSSFCell cell = row.createCell(c);

                cell.setCellValue("Cell " + r + " " + c);
            }
        }

        FileOutputStream fileOut = new FileOutputStream(excelFileName);

        //write this workbook to an Outputstream.
        wb.write(fileOut);
        fileOut.flush();
        fileOut.close();
    }

}
