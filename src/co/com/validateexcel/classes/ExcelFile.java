/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package co.com.validateexcel.classes;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author EQUIPO
 * 
 * Esta clase permite la manipulacion de un archivo excel
 */
public class ExcelFile {

    public ExcelFile() {
    }
    
    /**
     * Metodo para leer un archivo excel
     * @param pathFile, ruta del archivo a leer
     * @return workBook || null, libro excel o nulo si algo falla
     */
    public XSSFWorkbook read(String pathFile){
        try {
            // Se lee el archivo
            File excelFile = new File(pathFile);
            // Se procesa el archivo
            InputStream excelStream = new FileInputStream(excelFile);
            // Se crea el libro de trabajo a partir de la lectura del archivo
            XSSFWorkbook workBook = new XSSFWorkbook(excelStream);
            // Se retorna el libro de excel
            return workBook;
        } catch (FileNotFoundException ex) {
            System.err.println("No se encontró el fichero (ERROR: 201): " + ex);
        } catch (IOException ex) {
            System.err.println("Error al procesar el fichero (ERROR: 202): " + ex);
        }
        return null;
    }
    
    /**
     * Metodo que imprime un nmero de filas
     * @param workBook, libro de excel
     * @param numSheet, posicion de la hoja en el libro
     */
    public void printSheet(XSSFWorkbook workBook,int numSheet){
        System.out.println("#!#!#! Imprimiendo hoja numero: " + numSheet + " #!#!#!");
        XSSFSheet sheet = workBook.getSheetAt(numSheet);
        int totalRows = sheet.getLastRowNum();
        this.printRows(sheet, totalRows);
    }
    
    /**
     * Metodo para imprimir filas de una hoja
     * @param sheet, hoja del libro
     * @param numRowsPrint, numero de filas a imprimir
     */
    private void printRows(XSSFSheet sheet, int numRowsPrint){
        System.out.println("Numero de filas a imprimir: " + numRowsPrint);
        // ciclo que recorre todas las filas de la hoja
        for (int numRow = 0; numRow <= numRowsPrint; numRow++) {
            XSSFRow xssfRow = sheet.getRow(numRow);
            if (xssfRow == null){
                break;
            }else{
                System.out.print("Fila " + numRow + " -> ");
                for (int numCol = 0; numCol < xssfRow.getLastCellNum(); numCol++) {
                    String cellValue;
                    cellValue = xssfRow.getCell(numCol) == null?"":
                            (xssfRow.getCell(numCol).getCellType() == CellType.STRING)?xssfRow.getCell(numCol).getStringCellValue():
                            (xssfRow.getCell(numCol).getCellType() == CellType.NUMERIC)?"" + xssfRow.getCell(numCol).getNumericCellValue():
                            (xssfRow.getCell(numCol).getCellType() == CellType.BOOLEAN)?"" + xssfRow.getCell(numCol).getBooleanCellValue():
                            (xssfRow.getCell(numCol).getCellType() == CellType.BLANK)?"$$$BLANK$$$":
                            (xssfRow.getCell(numCol).getCellType() == CellType.FORMULA)?"$$$FORMULA$$$":
                            (xssfRow.getCell(numCol).getCellType() == CellType.ERROR)?"$$$ERROR$$$":
                            (xssfRow.getCell(numCol).getCellType() == CellType._NONE)?"$$$ERROR$$$":"$$$NONE$$$";                       
                    System.out.print("[Columna " + numCol + ": " + cellValue + "] ");
                }
                System.out.println("");
            }
        }
    }
    
    
    
//    System.out.println("\n num:" + CellType.NUMERIC);
//    System.out.println("str:" + CellType.STRING);
//    System.out.println("blank:" + CellType.BLANK);
//    System.out.println("bool:" + CellType.BOOLEAN);
//    System.out.println("err:" + CellType.ERROR);
//    System.out.println("formula:" + CellType.FORMULA);
//    System.out.println("formula:" + CellType._NONE);
}
