package sheet_snippets;

import java.io.File;
import java.io.FileInputStream;
import java.net.URL;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.*;

public class ReadExcelSheet {

    public static void main(String args[]){
        String excelFileName = "myfile.xlsx";
        URL excelFilePath = ReadExcelSheet.class.getResource(excelFileName);
        File excelFile_ = new File(excelFilePath.getFile());
        System.out.println("Starting process....");
        readExcel(excelFile_);

    }

    private static void readExcel(File excelFile_){
        try {

            FileInputStream excelFile = new FileInputStream(excelFile_);
            Workbook workbook = WorkbookFactory.create(excelFile);
            int noOfSheets = workbook.getNumberOfSheets();
            for (int i=0; i < noOfSheets; i++){
                Sheet sheet = workbook.getSheetAt(i);
                Iterator<Row> sheetRowIterator = sheet.iterator();
                while (sheetRowIterator.hasNext()){
                    String rowString = "";
                    Row row = sheetRowIterator.next();
                    Iterator<Cell> cellIterator = row.cellIterator();

                    while(cellIterator.hasNext()){
                        Cell cell = cellIterator.next();
                        rowString = rowString + cell.getStringCellValue() + " ";
                    }

                    System.out.println(rowString);
                }
            }
            System.out.println();
        }

        catch(Exception exc){
            System.out.println(exc.getMessage());
            exc.printStackTrace();
        }
    }
}
