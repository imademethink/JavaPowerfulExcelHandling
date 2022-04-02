package com.PowerfulExcelWriting;

public class ExcelReporting01 {


    // Brief about apache-poi libraries
    // HSSF (Horrible Spreadsheet Format) (for XLS)
    // XSSF (XML Spreadsheet Format) (for XLSX)
    // workbook, sheet, row, cell


    public static void main(String[] args) throws Exception{

        String baseFolder = System.getProperty("user.dir") + "\\reports";

        // XL creation (with 1 worksheet, name : Sheet1 (default))
        String ex01_DefaultXLPath = baseFolder + "\\example01\\DefaultExcel.xlsx";
        Util01_XLCreation.CreateXL(ex01_DefaultXLPath,null,true);
        System.out.println("Creating XL at "+ex01_DefaultXLPath);
        System.out.println();

        // XL copy
        String ex01_CopyXLPath    = baseFolder + "\\example01\\CopyExcel.xlsx";
        Util01_XLCreation.CopyXL(ex01_DefaultXLPath,ex01_CopyXLPath);
        System.out.println("Copying XL from "+ex01_DefaultXLPath);
        System.out.println("Copying XL to   "+ex01_CopyXLPath);
        System.out.println();

        // Add blank sheet
        Util01_XLCreation.AddBlankWorksheetAfterIndex(ex01_DefaultXLPath,1,"Sheet2"); // index starts with 0
        Util01_XLCreation.AddBlankWorksheetAfterIndex(ex01_DefaultXLPath,2,"Sheet3"); // index starts with 0
        Util01_XLCreation.AddBlankWorksheetAfterIndex(ex01_DefaultXLPath,2,"Sheet3New"); // index starts with 0
        Util01_XLCreation.AddBlankWorksheetAfterIndex(ex01_DefaultXLPath,3,"Sheet4"); // index starts with 0
        System.out.println("Adding blank sheet into "+ex01_DefaultXLPath);
        System.out.println("                   at index "+1);
        System.out.println();

        // Clone/ Copy worksheet in same XL
        Util01_XLCreation.CloneWorksheet(ex01_DefaultXLPath, 1, 2);
        System.out.println("Cloning worksheet from index 1 to index 2 for "+ex01_DefaultXLPath);
        System.out.println();

        // Copy worksheet in another XL (Currently with limited Cell Style copying)
        ex01_DefaultXLPath          = baseFolder + "\\example01\\Input.xlsx";
        String ex01_CopyToXLPath    = baseFolder + "\\example01\\CopyToExcel.xlsx";
        Util01_XLCreation.CreateXL(ex01_CopyToXLPath,null,true);
        Util01_XLCreation.CopyWorksheetFromTo(ex01_DefaultXLPath, 0,
                                              ex01_CopyToXLPath, 0);
        System.out.println("Copying worksheet from path  "+ex01_DefaultXLPath);
        System.out.println("                  from index "+0);
        System.out.println("Copying worksheet to path    "+ex01_CopyToXLPath);
        System.out.println("                  to index   "+0);
        System.out.println();

        // Delete XL
        ex01_DefaultXLPath = baseFolder + "\\example01\\DefaultExcel2.xlsx";
        Util01_XLCreation.CreateXL(ex01_DefaultXLPath,null,true);
        System.out.println("Creating XL at path "+ex01_DefaultXLPath);
        Util01_XLCreation.DeleteXL(ex01_DefaultXLPath);
        System.out.println("Deleting XL at path "+ex01_DefaultXLPath);

    }

}

