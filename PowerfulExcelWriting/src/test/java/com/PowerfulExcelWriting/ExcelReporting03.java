package com.PowerfulExcelWriting;
import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
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

public class ExcelReporting03 {

    public static void main(String[] args) throws Exception{

        String basePath   = System.getProperty("user.dir");
        String sExcelPath = basePath + "\\cell_formatting\\simpleformatting.xlsx";

        Util01_XLCreation.CreateXL(sExcelPath,null,true);
        System.out.println("Creating XL at "+sExcelPath);


        // Set cell value on new sheet on new cell
        String aString = "July";
        System.out.println("\nUtilLog: Cell value set to String " + aString);
        Util02_XLReading.mExcel_SetCellValue(sExcelPath,"Sheet1",2,1,aString);

        // Merge few cells
        System.out.println("\nUtilLog: Merging Cells");
        int firstRow = 2; int lastRow=9; int firstCol=1; int lastCol=3;
        int nMergedRegion1 = Util03_XLFormatting.mMergeFollowingRegion(sExcelPath, "Sheet1",
                                                        firstRow,lastRow,firstCol,lastCol);
        int nMergedRegion2 = Util03_XLFormatting.mMergeFollowingRegion(sExcelPath, "Sheet1",
                firstRow+3,lastRow+3,firstCol+4,lastCol+3);

        // Get total merged regions
        System.out.println("\nUtilLog: Total merged regions currently are : " +
                Util03_XLFormatting.mGetTotalMergedRegion(sExcelPath, "Sheet1")
        );

        // Get Merged region info
        Util03_XLFormatting.mGetMergedRegionInfo(sExcelPath, "Sheet1", (nMergedRegion1 -1) );
        // As 2nd region has no cell with any value assigned so can not get full region info
        // Util03_XLFormatting.mGetMergedRegionInfo(sExcelPath, "Sheet1", (nMergedRegion2 -1) );

        // To insert value into merged region, set it to top-left cell
        String aStringNew = "November";
        System.out.println("\nUtilLog: Cell value set to String " + aStringNew);
        Util02_XLReading.mExcel_SetCellValue(sExcelPath,"Sheet1",
                            firstRow+3,firstCol+4,aStringNew);
//        // Unmerge region
//        System.out.println("\nUtilLog: UnMerging Cells / Region");
//        Util03_XLFormatting.mUnMergeFollowingRegion(sExcelPath, "Sheet1",(nMergedRegion2 - 1));
//        Util03_XLFormatting.mUnMergeFollowingRegion(sExcelPath, "Sheet1",(nMergedRegion1 - 1));





//        int aInt = 7000;
//        System.out.println("\nUtilLog: Cell value set to Integer " + aInt);
//        Util02_XLReading.mExcel_SetCellValue(sExcelPath,"Sheet1",5,5,aInt);
//
//        float aFloat = -8.501f;
//        System.out.println("\nUtilLog: Cell value set to Float " + aFloat);
//        Util02_XLReading.mExcel_SetCellValue(sExcelPath,"Sheet1",7,5,aFloat);
//
//        boolean aBoolean = false;
//        System.out.println("\nUtilLog: Cell value set to Boolean " + aBoolean);
//        Util02_XLReading.mExcel_SetCellValue(sExcelPath,"Sheet1",8,5,aBoolean);



    }

}
