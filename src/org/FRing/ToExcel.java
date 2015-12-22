package org.FRing;
import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.InputStreamReader;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ToExcel {

  public static void convertCSVToExcel(String readinFile, String outputFile) {
    try {
      @SuppressWarnings("resource")
      BufferedReader readTxt = new BufferedReader(new InputStreamReader(new FileInputStream(readinFile), "UTF-8"));
      String inStr = "";
   
      HSSFWorkbook writeWorkbook = new HSSFWorkbook();
      HSSFSheet targetSheet = writeWorkbook.createSheet();
      HSSFRow targetRow;
      HSSFCell targetCell;
      int rowIndex = 0;
   
      while ((inStr = readTxt.readLine())!=null) { 
        targetRow = targetSheet.createRow(rowIndex);
        String str[] = inStr.split(",");
        for (int i = 0; i < str.length; i++) {
          targetCell = targetRow.createCell(i);
          targetCell.setCellValue(str[i]); 
        }
        rowIndex ++;
      }
   
      FileOutputStream outputExcel = new FileOutputStream(outputFile);
      writeWorkbook.write(outputExcel);
      outputExcel.flush();
      outputExcel.close();
    }
    catch (Exception e) {
      e.printStackTrace();
    }
  }
}