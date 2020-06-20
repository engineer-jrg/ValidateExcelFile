/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package co.com.validateexcel.main;

import co.com.validateexcel.classes.ExcelFile;
import co.com.validateexcel.classes.ValidateExcel;
import java.util.ArrayList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author EQUIPO
 */
public class main {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        
        // varibles para el archivo excel
        String filePath = "C:\\Users\\EQUIPO\\Documents\\Aos\\Bancolombia\\Test\\Validar-Excel\\";
        // String fileName = "BASILEA_FORWARD_31122019.xlsx";
        String fileName = "data.xlsx";
        // String sheet = "31122019";
        int numSheet = 0;
        
        // Se crea un objeto de la clase que lee el archivo
        ExcelFile excelFile = new ExcelFile();
        // Se lee el archivo
        XSSFWorkbook workBook = excelFile.read(filePath + fileName);
        // Se verifica la correcta lectura del archivo
        if(workBook == null){
            System.err.println("Fallo la lectura del libro (ERROR: 101)");
        } else {
            // la lectura del libro ocurrio sin problemas
            // imprimir el contenido de una hoja
            excelFile.printSheet(workBook, numSheet);
            
            // validaciones { {}, {}, {}, {} }
            String [][] validaciones = {
                {"date"}, // columna 1
                {"numeric"}, // columna 2
                {}, // columna 3
                {"string","noOnlyNumeric","strlength-14"} // columna 4
            };
            
            ValidateExcel validateExcel = new ValidateExcel();
            int ignoreRows = 1; // filas para ignorar
            
            // ejecutando la validacion
            ArrayList<String> errors = validateExcel.validate(workBook, numSheet, ignoreRows, validaciones);
            
            System.out.println("\n");
            System.out.println("#!#!#! Resultados de la validacion #!#!#!");
            System.out.println("");
            System.out.println("* numero de errores: " + errors.size());
            System.out.println("");
            for (String error : errors) {
                System.out.println("Error #1: " + error);
            }
            
            // cargar en la DB
            // out(fila) || error bull copy
        }
    }
    
}
