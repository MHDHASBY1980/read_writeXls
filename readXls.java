/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package read_writeXls;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;

/**
 *
 * @author USER
 */
public class readXls {
    public static void main(String[] args) throws IOException {
        readFromExcel("D://readXls.xls");
    }
    public static void readFromExcel(String urlexcel) throws FileNotFoundException, IOException{
        HSSFWorkbook myexcel = new HSSFWorkbook(new FileInputStream(urlexcel));
        HSSFSheet myexcelSheet = myexcel.getSheet("training");
        FormulaEvaluator formulaEv = myexcel.getCreationHelper().createFormulaEvaluator();
        
        for(Row row: myexcelSheet){
           for(Cell cell:row){
               switch(formulaEv.evaluateInCell(cell).getCellType()){
                   case Cell.CELL_TYPE_NUMERIC:
                       System.out.print(cell.getNumericCellValue() + "\t\t");
                       break;
                   case Cell.CELL_TYPE_STRING:
//                       if(cell.getColumnIndex()==0){
                           System.out.print(cell.getStringCellValue() + "\t\t");
//                       }
//                       else if(cell.getColumnIndex()==1){
//                           System.out.println(cell.getStringCellValue() + "\t\t");
//                       }
//                       else if(cell.getColumnIndex()==2){
//                           System.out.println(cell.getStringCellValue() + "\t\t"); 
//                       }
                       break;
               }
           } 
            System.out.println("");
            myexcel.close();
        }
        
        
    }
    
}
