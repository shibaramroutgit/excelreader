package org.example;

import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

import java.io.IOException;
import java.util.List;

public class TestCase1 extends TestBase
{
    public String excelPath = "src/test/java/org/example/excelFiles/UserData.xlsx";
    public ExcelUtilities excelUtilities;
    public String sheetName ="IndiaUser";

    @BeforeMethod
    public void runBeforeMethod() throws IOException {
        excelUtilities= new ExcelUtilities(excelPath,sheetName);
    }

    @Test
    public void testCase1() throws IOException {
        int x = excelUtilities.getRowCountOfSheet();
        int y = excelUtilities.getColumnCountOfSheet();
        String val = excelUtilities.getSpecificCellData(2,3);
        List<String> rowDataList = excelUtilities.getAllCellDataOfARow(2);
        List<String> allDataList = excelUtilities.getAllDataFromSheet();
        System.out.println(x);
        System.out.println(y);
        System.out.println(val);
        System.out.println(rowDataList.toString());
        System.out.println(allDataList.toString());

    }
}
