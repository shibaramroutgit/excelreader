package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelUtilities {
    public FileInputStream fileInputStream;
    public XSSFWorkbook xssfWorkbookworkbook;
    public XSSFSheet xssfSheetsheet;
    public XSSFRow xssfRowrow;
    public XSSFCell xssfCellcell;
    public  ExcelUtilities(String excelPath,String nameOfSheet) throws IOException {
        fileInputStream = new FileInputStream(excelPath);
        xssfWorkbookworkbook = new XSSFWorkbook(fileInputStream);
        xssfSheetsheet = xssfWorkbookworkbook.getSheet(nameOfSheet);
    }

    public int getRowCountOfSheet()
    {
       return  xssfSheetsheet.getLastRowNum();
    }

    public int getColumnCountOfSheet()
    {
        return  xssfSheetsheet.getRow(0).getLastCellNum();
    }

    public String getSpecificCellData(int rowNo, int cellNo)
    {
        String returnCellVal = null;
       Cell cell =  xssfSheetsheet.getRow(rowNo).getCell(cellNo);
       if(cell.getCellType()== CellType.STRING)
       {
           returnCellVal= cell.getStringCellValue();
       } else if (cell.getCellType()==CellType.NUMERIC) {
           returnCellVal= String.valueOf(cell.getNumericCellValue());
       }
       return returnCellVal;
    }

    public List<String> getAllCellDataOfARow(int rowNo)
    {
        List<String> rowDataList = new ArrayList<String>();

       Row row =  xssfSheetsheet.getRow(rowNo);
       //int lastColumnNum = xssfSheetsheet.getRow(rowNo).getLastCellNum();
       for (Cell cell:row)
       {
           if(cell.getCellType()==CellType.STRING) {
               rowDataList.add(cell.getStringCellValue());
           } else if (cell.getCellType()==CellType.NUMERIC) {
               rowDataList.add(String.valueOf(cell.getNumericCellValue()));
           }
       }
       return rowDataList;
    }

    public List<String> getAllDataFromSheet()
    {
        List<String> allDataOfSheet = new ArrayList<String>();
        for(int i=1;i<=xssfSheetsheet.getLastRowNum();i++)
        {
            Row row = xssfSheetsheet.getRow(i);
            for (Cell cell:row)
            {
                if(cell.getCellType()==CellType.STRING) {
                    allDataOfSheet.add(cell.getStringCellValue());
                } else if (cell.getCellType()==CellType.NUMERIC) {
                    allDataOfSheet.add(String.valueOf(cell.getNumericCellValue()));
                }
            }
        }

        return allDataOfSheet;
    }


}
