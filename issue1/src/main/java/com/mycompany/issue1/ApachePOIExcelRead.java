/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.issue1;

import java.io.BufferedWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.Writer;
import java.util.Iterator;

public class ApachePOIExcelRead {


    private static final String FILE_NAME = "/C:\\Users\\Naim Saleh\\Downloads\\Documents\\list name.xlsx";
    static boolean dash = true;
    public static void main(String[] args) throws IOException {
        Writer w = null;
        File file = new File ("C:\\Users\\Naim Saleh\\231919-STIW3054-A172-A1.wiki\\Result.md");
        w = new BufferedWriter(new FileWriter(file));

        try {

            FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();

            while (iterator.hasNext()) {

                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();

                while (cellIterator.hasNext()) {

                    Cell currentCell = cellIterator.next();
                    //getCellTypeEnum shown as deprecated for version 3.15
                    //getCellTypeEnum ill be renamed to getCellType starting from version 4.0
                    if (currentCell.getCellTypeEnum() == CellType.STRING) {
                        System.out.print(currentCell.getStringCellValue() + "  |  ");
                        w.write(currentCell.getStringCellValue() + "  |  ");
                    } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                        System.out.print(currentCell.getNumericCellValue() + "  |  ");
                        w.write(currentCell.getNumericCellValue() + "  |  ");
                    }

                }
                System.out.println();
                w.write("\n");
                if (dash == true){
                    System.out.println("|-|-|-|-| \n");
                    w.write("|-|-|-|-| \n");
                    dash = false;
                }

            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        
        try {
            if(w != null){
                w.close();
            }
        }
        catch(IOException e){
            e.printStackTrace();
        }
            
            

    }
}
