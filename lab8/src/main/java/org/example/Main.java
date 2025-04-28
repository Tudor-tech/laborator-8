package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class Main {
    public static void main(String[] args) {
        FileInputStream file;
        XSSFSheet sheet;
        try {
            file = new FileInputStream(new File("TestExcel.xlsx"));
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            sheet = workbook.getSheetAt(0);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                switch (cell.getCellType()) {
                    case NUMERIC:
                        System.out.println(cell.getNumericCellValue());
                        break;
                    case STRING:
                        System.out.println(cell.getStringCellValue());
                        break;
                    case FORMULA:
                        System.out.println(cell.getCellFormula());

                }
            }
            System.out.println("");
        }

        XSSFWorkbook worklook = new XSSFWorkbook();
        XSSFSheet sheet1 = worklook.createSheet("TestExcel");

        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("5", new Object[] {"Slavu", "Tudor", "20"});

        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset) {

            Row row = sheet1.createRow(rownum++);
            Object [] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr)
            {
                Cell cell = row.createCell(cellnum++);
                if(obj instanceof String)
                    cell.setCellValue((String)obj);
                else if(obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
            }
        }
        try {
            FileOutputStream out = new FileOutputStream(new File("altFisier.xlsx"));
            worklook.write(out);
            out.close();
            System.out.println("ai scris in fisier");
        }
        catch (IOException e) {
            e.printStackTrace();
        }
    }
}