package com.PowerfulExcelWriting;
import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.junit.Assert;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

public class ExcelReporting02 {

    public static void main(String[] args) throws Exception {

        String basePath   = System.getProperty("user.dir");
        String sExcelPath = basePath + "\\inputXL\\file_excel_simple.xlsx";


        // Reading cell value
        XSSFCell oCell_number = Util02_XLReading.mExcel_GetCellValue(sExcelPath,"emp_data",1,0);
        System.out.println("\nUtilLog: Cell value received " + oCell_number);
        System.out.println("UtilLog: Actual cell type " + Util02_XLReading.mGetExactCellType(oCell_number));
        int n = Integer.parseInt(new DataFormatter().formatCellValue(oCell_number));
        System.out.println("UtilLog: Converted cell value integer " + n);

        XSSFCell oCell_string = Util02_XLReading.mExcel_GetCellValue(sExcelPath,"emp_data",1,1);
        System.out.println("\nUtilLog: Cell value received " + oCell_string);
        System.out.println("UtilLog: Actual cell type " + Util02_XLReading.mGetExactCellType(oCell_string));
        String s = oCell_string.getStringCellValue();
        System.out.println("UtilLog: Converted cell value string " + s);

        XSSFCell oCell_number2 = Util02_XLReading.mExcel_GetCellValue(sExcelPath,"emp_data",1,3);
        System.out.println("\nUtilLog: Cell value received " + oCell_number2);
        System.out.println("UtilLog: Actual cell type " + Util02_XLReading.mGetExactCellType(oCell_number2));
        float f = Float.parseFloat(oCell_number2.toString());
        System.out.println("UtilLog: Converted cell value float " + f);

        XSSFCell oCell_boolean = Util02_XLReading.mExcel_GetCellValue(sExcelPath,"emp_data",1,5);
        System.out.println("\nUtilLog: Cell value received " + oCell_boolean);
        System.out.println("UtilLog: Actual cell type " + Util02_XLReading.mGetExactCellType(oCell_boolean));
        boolean b = Boolean.parseBoolean(oCell_boolean.toString());
        System.out.println("UtilLog: Converted cell value boolean " + b);

        XSSFCell oCell_blank = Util02_XLReading.mExcel_GetCellValue(sExcelPath,"emp_data",1,6);
        System.out.println("\nUtilLog: Cell value received " + oCell_blank);
        System.out.println("UtilLog: Actual cell type " + Util02_XLReading.mGetExactCellType(oCell_blank));
        System.out.println("UtilLog: Converted cell value blank or just white space (s) "+ oCell_blank);



//
//        // Set cell value on existing sheet, existing cell
//        int aInt = 7000;
//        System.out.println("\nUtilLog: Cell value set to Integer " + aInt);
//        Util02_XLReading.mExcel_SetCellValue(sExcelPath,"emp_data",1,0,aInt);
//
//        String aString = "July";
//        System.out.println("\nUtilLog: Cell value set to String " + aString);
//        Util02_XLReading.mExcel_SetCellValue(sExcelPath,"emp_data",1,1,aString);
//
//        float aFloat = -8.501f;
//        System.out.println("\nUtilLog: Cell value set to Float " + aFloat);
//        Util02_XLReading.mExcel_SetCellValue(sExcelPath,"emp_data",1,3,aFloat);
//
//        boolean aBoolean = false;
//        System.out.println("\nUtilLog: Cell value set to Boolean " + aBoolean);
//        Util02_XLReading.mExcel_SetCellValue(sExcelPath,"emp_data",1,5,aBoolean);



        // Set cell value on new sheet on new cell
        int aInt = 7000;
        System.out.println("\nUtilLog: Cell value set to Integer " + aInt);
        Util02_XLReading.mExcel_SetCellValue(sExcelPath,"emp_data_new",5,5,aInt);

        String aString = "July";
        System.out.println("\nUtilLog: Cell value set to String " + aString);
        Util02_XLReading.mExcel_SetCellValue(sExcelPath,"emp_data_new",6,5,aString);

        float aFloat = -8.501f;
        System.out.println("\nUtilLog: Cell value set to Float " + aFloat);
        Util02_XLReading.mExcel_SetCellValue(sExcelPath,"emp_data_new",7,5,aFloat);

        boolean aBoolean = false;
        System.out.println("\nUtilLog: Cell value set to Boolean " + aBoolean);
        Util02_XLReading.mExcel_SetCellValue(sExcelPath,"emp_data_new",8,5,aBoolean);





//        // Exploring cell styling if required
//        // Open the reference/ backup excel first
//        Util02_XLReading.mExcel_CellStyle_Explanation();


//        // Shifting row(s), column(s)
//        // Open the reference/ backup excel first
//        Util02_XLReading.mExcel_Shifting_Operation();






//        String sExcelPath_employee = basePath + "\\inputXL\\file_excel_data_employee.xlsx";

//        // read header and all data rows as ArrayList of ArrayList
//        ArrayList<ArrayList<Object>> arylst2D = Util02_XLReading.mExcel_ReadData2DAryList(sExcelPath_employee,"data");

//        // read header and all data rows as 2D array of Objects
//        sExcelPath_employee = basePath + "\\inputXL\\file_excel_data_employee.xlsx";
//        Object[][] ary2D = Util02_XLReading.mExcel_ReadData2DAry(sExcelPath_employee,"data");


//        // read header and particular data row
//        String sExcelPath_banking = basePath + "\\inputXL\\file_excel_data_banking.xlsx";
//        HashMap<String, Object> hmBankingDataRow = Util02_XLReading.mExcel_ReadDataRow(sExcelPath_banking, "data",1);
//        System.out.println("\nReading HashMap data row (key-value) for banking");
//        Util02_XLReading.mPrintDataRow(hmBankingDataRow);



//        // read header and particular data row for specific unique id
//        String sExcelPath_banking2 = basePath + "\\inputXL\\file_excel_data_banking.xlsx";
//        String sSpecificBankId     = "bank_34099";
//        HashMap<String, Object> hmBankingDataRowSpecific = Util02_XLReading.mExcel_ReadDataRow_UniqueId(
//                sExcelPath_banking2, "data",sSpecificBankId);
//        System.out.println("\nReading HashMap data row for banking for specific bank id " + sSpecificBankId);
//        Util02_XLReading.mPrintDataRow(hmBankingDataRowSpecific);




//        // read cell value for particular data row with specific unique id (in 0th coln) for specific column
//        String sExcelPath_ecommerce2 = basePath + "\\inputXL\\file_excel_data_ecommerce.xlsx";
//        String sSpecificCustId       = "ecom_34099";
//        String sSpecificBankColnName = "orderid";
//        XSSFCell oCellBanking        = Util02_XLReading.mExcel_ReadCell_UniqueId_SpecificCol(
//                sExcelPath_ecommerce2, "data",sSpecificCustId, sSpecificBankColnName);
//        System.out.println("\nReading Cell value for banking for" +
//                " specific customer id " + sSpecificCustId +
//                " specific column name " + sSpecificBankColnName +
//                " is " + oCellBanking);


//        // Dump database result set to excel
//        Util02_XLReading.mExcel_DumpResultSet_to_Excel();


//        // Convert excel to pdf (with correct orientation
//        Util02_XLReading.mExcel_ConvertTestCaseExcelToPdf();


//
//        // Process Excel like an SQL SELECT, UPDATE…. query
//        String sExcelPath_banking3 = basePath + "\\inputXL\\file_excel_data_banking.xlsx";
//        String sSheetName          = "data";
//        String sSpecificBankId     = "bank_34099";
//        String sSelectQuery        = "select * from  " + sSheetName;
//        String sSelectQuery2       = "select * from  " + sSheetName + " where cust_id='" + sSpecificBankId + "'";
//        Recordset oRecordset       = Util02_XLReading.mExcel_ReadDataRow_UniqueId_SelectQuery(
//                                                                          sExcelPath_banking3, sSelectQuery2);



    }


}
