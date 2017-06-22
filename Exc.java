package com.test.local;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

public class Exc {

    private static final String FILE_NAME = "E:/abc/abc.xls";

    public static void main(String[] args) throws InvalidFormatException {

        Map<String, Double> map = new HashMap<>(); 
        try {
            FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
            Workbook workbook = new HSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();

            while (iterator.hasNext()) {

                Row currentRow = iterator.next();
                if(currentRow.getCell(4) == null)
                {
                    Double aDouble = map.get(currentRow.getCell(1).getStringCellValue());
                    if(aDouble != null)
                    {
                        map.put(currentRow.getCell(1).getStringCellValue(),aDouble + currentRow.getCell(2).getNumericCellValue());
                    }
                    else
                    {
                        map.put(currentRow.getCell(1).getStringCellValue(),currentRow.getCell(2).getNumericCellValue());
                    }
                    System.out.println(currentRow.getCell(0) + "\t" + currentRow.getCell(1) + "\t" + currentRow.getCell(2) + "\t" + currentRow.getCell(3));
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("--HOLD--");
        Set<Map.Entry<String, Double>> entries = map.entrySet();
        for(Map.Entry<String, Double> entry:entries)
        {
            System.out.println(entry.getKey() + "\t" + entry.getValue());
        }

    }
}

/*
ignore 1st few rows

                if (currentCell.getCellTypeEnum() == CellType.STRING) {
                        System.out.print(currentCell.getStringCellValue() + "--");
                        } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                        System.out.print(currentCell.getNumericCellValue() + "--");
                        }
*/
