import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;

/**
 * Created by sdlds on 2017/4/19.
 */
public class WriteToExcel {
    private static String per = "%";
    private static String add = "+";
    private static String mis = "-";
    private static String milton = "万吨";
    private static String bilyuan = "亿元";
    private static String hu = "户";

    public static void writeToExcel2006( ArrayList<String> temp, String title){
        Workbook wb = null;
        try {
            FileInputStream fileInputStream = new FileInputStream("file/" + title + ".xlsx");
            wb = WorkbookFactory.create(fileInputStream);
        }catch (Exception e){
            e.printStackTrace();
        }
        Sheet sheet = null;
        sheet = wb.getSheet("综合");
        sheet.getRow(1).createCell(2).setCellValue(temp.get(475));
        sheet.getRow(1).createCell(3).setCellValue("");
        sheet.getRow(2).createCell(2).setCellValue("");
        sheet.getRow(2).createCell(3).setCellValue("");
        sheet.getRow(3).createCell(2).setCellValue(temp.get(477));
        sheet.getRow(3).createCell(3).setCellValue("");
        sheet.getRow(4).createCell(2).setCellValue(temp.get(478));
        sheet.getRow(4).createCell(3).setCellValue("");
        sheet.getRow(5).createCell(2).setCellValue(temp.get(470));
        sheet.getRow(5).createCell(3).setCellValue("");
        sheet.getRow(6).createCell(2).setCellValue(temp.get(472));
        sheet.getRow(6).createCell(3).setCellValue("");
        sheet.getRow(7).createCell(2).setCellValue("");
        sheet.getRow(7).createCell(3).setCellValue("");
        sheet.getRow(8).createCell(2).setCellValue(temp.get(482) + per);

        sheet.getRow(11).createCell(2).setCellValue(temp.get(6));
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(7));
        sheet.getRow(12).createCell(2).setCellValue("");
        sheet.getRow(12).createCell(3).setCellValue("");
        sheet.getRow(13).createCell(2).setCellValue(temp.get(27));
        sheet.getRow(13).createCell(3).setCellValue(add + temp.get(29));
        sheet.getRow(14).createCell(2).setCellValue(temp.get(9));
        sheet.getRow(14).createCell(3).setCellValue(add + temp.get(10));
        sheet.getRow(15).createCell(2).setCellValue(temp.get(11));
        sheet.getRow(15).createCell(3).setCellValue(add + temp.get(12));
        sheet.getRow(16).createCell(2).setCellValue(temp.get(13));
        sheet.getRow(16).createCell(3).setCellValue(add + temp.get(14));

        sheet.getRow(20).createCell(2).setCellValue("");
        sheet.getRow(21).createCell(2).setCellValue("");



        sheet = wb.getSheet("农业");
        sheet.getRow(1).createCell(2).setCellValue("");
        sheet.getRow(1).createCell(3).setCellValue("");
        sheet.getRow(2).createCell(2).setCellValue(temp.get(77));
        sheet.getRow(2).createCell(3).setCellValue(add + temp.get(79));
        sheet.getRow(3).createCell(2).setCellValue(temp.get(80));
        sheet.getRow(3).createCell(3).setCellValue(mis + temp.get(81));
        sheet.getRow(4).createCell(2).setCellValue(temp.get(82));
        sheet.getRow(4).createCell(3).setCellValue(add + temp.get(83));

        sheet.getRow(10).createCell(2).setCellValue(temp.get(84) + milton);
        sheet.getRow(11).createCell(2).setCellValue("");

        sheet.getRow(20).createCell(2).setCellValue(temp.get(86) + milton);


        sheet = wb.getSheet("工业");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(141) + bilyuan);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(151) + hu);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(187) + bilyuan);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(185) + bilyuan);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(183) + bilyuan);


        try {
            String excelPath = "file/" + title + ".xlsx";
            FileOutputStream fileOutputStream = new FileOutputStream(excelPath);
            wb.write(fileOutputStream);
            fileOutputStream.flush();
            fileOutputStream.close();
            System.out.println("Success");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void writeToExcel2008(Workbook wb, ArrayList<String> temp){

    }

    public static void writeToExcel2009(Workbook wb, ArrayList<String> temp){

    }

    public static void writeToExcel2010(Workbook wb, ArrayList<String> temp){

    }

    public static void writeToExcel2011(Workbook wb, ArrayList<String> temp){

    }

    public static void writeToExcel2012(Workbook wb, ArrayList<String> temp){

    }

    public static void writeToExcel2013(Workbook wb, ArrayList<String> temp){

    }

    public static void writeToExcel2014(Workbook wb, ArrayList<String> temp){

    }

    public static void writeToExcel2015(Workbook wb, ArrayList<String> temp){

    }

    public static void writeToExcel2016(Workbook wb, ArrayList<String> temp){

    }
}
