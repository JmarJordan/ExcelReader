package com.bib.readexcel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

public class Launcher {

    public static void main(String[] args) {
        try {
            //FileInputStream file = new FileInputStream("C:\\Users\\ef-jeymar\\Documents\\Test.xlsx");
            FileInputStream file = new FileInputStream("D:\\mytest.xlsx");

            Workbook workbook = new XSSFWorkbook(file);

            //Get first/desired sheet from the workbook
            Sheet sheet = workbook.getSheetAt(0);

            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    CellType type = cell.getCellType();
                    DataFormatter formatter = new DataFormatter();
                    String cellView = formatter.formatCellValue(cell);
                    switch (type){
                        case STRING:
                        case BOOLEAN:
                        case NUMERIC:
                            System.out.print(cellView + "\t");
                            break;
                    }


                }
                System.out.println(" ");
            }
            file.close();

        } catch (Exception e) {
            e.printStackTrace();


        }
    }
}
