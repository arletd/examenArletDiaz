/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package examen_arletdiazmendez;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import static org.apache.poi.ss.usermodel.CellType.BLANK;
import static org.apache.poi.ss.usermodel.CellType.NUMERIC;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 *
 * @author arlet
 */
public class Examen_ArletDiazMendez {
    
    
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException {
        String srcPath = "Mapeo Colaborativo datos.gob.mx __ bit.ly_rescateMX.xlsx";
        guardaNecesidades(srcPath);
    }
    
     public static void guardaNecesidades(String sourceBook) throws IOException{
        XSSFWorkbook origenDatos = new XSSFWorkbook(new FileInputStream(sourceBook));
        XSSFSheet centrosDeAcopio = origenDatos.getSheet("Centros de Acopio - Colaborativ");
        Iterator<Row> iterator = centrosDeAcopio.iterator();
        XSSFWorkbook newBook = new XSSFWorkbook();
        XSSFSheet sheet = newBook.createSheet("Necesidades");
        int rowCount = 0;
        Row nextRow = iterator.next();
        Row newRow = sheet.createRow(rowCount++);
        Cell cellId;
        Cell cellNom;
        Cell cellNecesidad;
        do{
            int columnCount = 0;
            String necesidades = nextRow.getCell(3).getStringCellValue();
            String [] necesidadesArray= necesidades.split(",");
            for(int i = 0; i < necesidadesArray.length; i++){
                if(nextRow.getCell(0).getCellType() == NUMERIC){
                    double id = nextRow.getCell(0).getNumericCellValue();
                    cellId = newRow.createCell(columnCount++);
                    cellId.setCellValue(id);
                }else{
                    String id = nextRow.getCell(0).getStringCellValue();
                    cellId = newRow.createCell(columnCount++);
                    cellId.setCellValue(id);
                }
                String nomCentro = nextRow.getCell(1).getStringCellValue();
                cellNom = newRow.createCell(columnCount++);
                cellNom.setCellValue(nomCentro);
                cellNecesidad = newRow.createCell(columnCount++);
                cellNecesidad.setCellValue(necesidadesArray[i]);
                newRow = sheet.createRow(rowCount++);
                columnCount=0;
                System.out.print(nomCentro + "\t");
                System.out.print(necesidadesArray[i] + "\t");
                System.out.println("");
            }
            nextRow = iterator.next();
            
        }while(iterator.hasNext() && nextRow.getCell(0).getCellType()!=BLANK); 
        sheet.autoSizeColumn(1);
        sheet.autoSizeColumn(2);
        try (FileOutputStream outputStream = new FileOutputStream("Centros de Acopio – Necesidades.xlsx")) {
            newBook.write(outputStream);
        }
        origenDatos.close();
   
    }    
}
