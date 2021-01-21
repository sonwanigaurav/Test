import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * This program illustrates how to update an existing Microsoft Excel document.
 * Append new rows to an existing sheet.
 */
public class ExcelUpdate {
    private static final String excelFilePath = "/Users/g9709/Sonwani/Code/ExcelUpdate/src/main/java/ExcelUpdate.xlsx";

    public static void main(String[] args) {
        try {
            //insertDataInExcel();
            updateExistingRow();

        } catch (Exception  ex) {
            ex.printStackTrace();
        }
    }

    private static void insertDataInExcel() throws IOException {
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
        Workbook workbook = WorkbookFactory.create(inputStream);

        Sheet sheet = workbook.getSheetAt(0);

        Object[][] bookData = {
                {"Trello The Passionate Programmer Gaurav treloo", "Chad Fowler Sonwani", 16},
                {"Trello Software Craftmanship", "Pete McBreen", 26},
                {"Trello The Art of Agile Development", "James Shore", 32},
                {"Trello Continuous Delivery", "Jez Humble", 41},
                {" Trello Data From trello Card", "Jez Humble", 41}
        };

        int rowCount = sheet.getLastRowNum();

        for (Object[] aBook : bookData) {
            Row row = sheet.createRow(++rowCount);

            int columnCount = 0;

            Cell cell = row.createCell(columnCount);
            cell.setCellValue(rowCount);

            for (Object field : aBook) {
                cell = row.createCell(++columnCount);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
            }

        }

        inputStream.close();

        FileOutputStream outputStream = new FileOutputStream(excelFilePath);
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
    }

    public static void updateExistingRow() throws IOException {
        FileInputStream fsIP= new FileInputStream(new File(excelFilePath)); //Read the spreadsheet that needs to be updated

        XSSFWorkbook wb = new XSSFWorkbook(fsIP); //Access the workbook

        XSSFSheet worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.

        Cell cell = worksheet.getRow(2).getCell(2);   // Access the second cell in second row to update the value

        cell.setCellValue("OverRide Last Name");  // Get current cell value value and overwrite the value

        fsIP.close(); //Close the InputStream

        FileOutputStream output_file =new FileOutputStream(new File(excelFilePath));  //Open FileOutputStream to write updates

        wb.write(output_file); //write changes

        output_file.close();  //close the stream
    }
}