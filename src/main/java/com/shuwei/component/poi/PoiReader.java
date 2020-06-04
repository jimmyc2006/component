package com.shuwei.component.poi;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;


public class PoiReader {
  public static void main(String[] args) throws IOException {
    InputStream inp = new FileInputStream("/Users/shuwei/Downloads/2.xls");
      Workbook wb = WorkbookFactory.create(inp);
      Sheet sheet = wb.getSheetAt(0);
      System.out.println(sheet.getSheetName());
      Row row = sheet.getRow(1);
      Cell cell = row.getCell(1);
      System.out.println(cell.getCellType());
      System.out.println(cell.getStringCellValue());
      // Write the output to a file
    }
}
