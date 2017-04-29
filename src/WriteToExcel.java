import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;

/**
 * Created by sdlds on 2017/4/19.
 */


/**
 * \d+.\d+|\d+
 */
public class WriteToExcel {
    private static String per = "%";
    private static String add = "+";
    private static String mis = "-";
    private static String milton = "万吨";
    private static String bilyuan = "亿元";
    private static String hu = "户";
    private static String bildollar = "亿美元";
    private static String km = "公里";
    private static String milhu = "万户";
    private static String suo = "所";
    private static String milpeo = "万人";

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
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(7) + per);
        sheet.getRow(12).createCell(2).setCellValue("");
        sheet.getRow(12).createCell(3).setCellValue("");
        sheet.getRow(13).createCell(2).setCellValue(temp.get(27));
        sheet.getRow(13).createCell(3).setCellValue(add + temp.get(29) + per);
        sheet.getRow(14).createCell(2).setCellValue(temp.get(9));
        sheet.getRow(14).createCell(3).setCellValue(add + temp.get(10) + per);
        sheet.getRow(15).createCell(2).setCellValue(temp.get(11));
        sheet.getRow(15).createCell(3).setCellValue(add + temp.get(12) + per);
        sheet.getRow(16).createCell(2).setCellValue(temp.get(13));
        sheet.getRow(16).createCell(3).setCellValue(add + temp.get(14) + per);

        sheet.getRow(20).createCell(2).setCellValue("");
        sheet.getRow(21).createCell(2).setCellValue("");

        sheet.getRow(30).createCell(2).setCellValue("");
        sheet.getRow(31).createCell(2).setCellValue("");
        sheet.getRow(32).createCell(2).setCellValue("");


        sheet.getRow(40).createCell(2).setCellValue("");


        sheet = wb.getSheet("农业");
        sheet.getRow(1).createCell(2).setCellValue("");
        sheet.getRow(1).createCell(3).setCellValue("");
        sheet.getRow(2).createCell(2).setCellValue(temp.get(77));
        sheet.getRow(2).createCell(3).setCellValue(add + temp.get(79) + per);
        sheet.getRow(3).createCell(2).setCellValue(temp.get(80));
        sheet.getRow(3).createCell(3).setCellValue(mis + temp.get(81) + per);
        sheet.getRow(4).createCell(2).setCellValue(temp.get(82));
        sheet.getRow(4).createCell(3).setCellValue(add + temp.get(83) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(84) + milton);
        sheet.getRow(11).createCell(2).setCellValue("");
        sheet.getRow(12).createCell(2).setCellValue("");

        sheet.getRow(20).createCell(2).setCellValue(temp.get(86) + milton);


        sheet = wb.getSheet("工业");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(141) + bilyuan);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(151) + hu);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(187) + bilyuan);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(185) + bilyuan);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(183) + bilyuan);


        sheet = wb.getSheet("投资");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(233) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(234) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(244) + bilyuan);
        sheet.getRow(10).createCell(3).setCellValue(add+ temp.get(245) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(246) + bilyuan);
        sheet.getRow(11).createCell(3).setCellValue(add+ temp.get(247) + per);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(248) + bilyuan);
        sheet.getRow(12).createCell(3).setCellValue(add+ temp.get(249) + per);


        sheet = wb.getSheet("贸易");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(253) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(254) + per);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(255) + bilyuan);
        sheet.getRow(1).createCell(3).setCellValue(add + temp.get(256) + per);
        sheet.getRow(2).createCell(2).setCellValue(temp.get(259) + bilyuan);
        sheet.getRow(2).createCell(3).setCellValue(add + temp.get(260) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(285) + bildollar);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(286) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(287) + bildollar);
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(288) + per);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(289) + bildollar);
        sheet.getRow(12).createCell(3).setCellValue(add + (Double.parseDouble(temp.get(290))*100) + per);


        sheet = wb.getSheet("交通");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(291) + km);
        sheet.getRow(1).createCell(2).setCellValue("");
        sheet.getRow(2).createCell(2).setCellValue("");
        sheet.getRow(3).createCell(2).setCellValue("");
        sheet.getRow(4).createCell(2).setCellValue("");
        sheet.getRow(5).createCell(2).setCellValue("");

        sheet.getRow(10).createCell(2).setCellValue(temp.get(306) + bilyuan);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(307) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(308) + milhu);
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(309) + per);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(310) + milhu);
        sheet.getRow(12).createCell(3).setCellValue(add + temp.get(311) + per);


        sheet = wb.getSheet("金融");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(324) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(325) + per);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(328) + bilyuan);
        sheet.getRow(1).createCell(3).setCellValue(add + temp.get(329) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(339) + bilyuan);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(340) + per);
        sheet.getRow(11).createCell(2).setCellValue("");
        sheet.getRow(11).createCell(3).setCellValue("");


        sheet = wb.getSheet("教育");
        sheet.getRow(0).createCell(1).setCellValue("");
        sheet.getRow(0).createCell(3).setCellValue("");
        sheet.getRow(0).createCell(5).setCellValue("");


        sheet.getRow(6).createCell(1).setCellValue(temp.get(345) + suo);
        sheet.getRow(6).createCell(2).setCellValue("");
        sheet.getRow(6).createCell(3).setCellValue(temp.get(346) + milpeo);
        sheet.getRow(7).createCell(1).setCellValue(temp.get(347) + suo);
        sheet.getRow(7).createCell(2).setCellValue("");
        sheet.getRow(7).createCell(3).setCellValue(temp.get(348) + milpeo);
        sheet.getRow(8).createCell(1).setCellValue(temp.get(349) + suo);
        sheet.getRow(8).createCell(2).setCellValue("");
        sheet.getRow(8).createCell(3).setCellValue(temp.get(350) + milpeo);
        sheet.getRow(9).createCell(1).setCellValue(temp.get(355) + suo);
        sheet.getRow(9).createCell(2).setCellValue("");
        sheet.getRow(9).createCell(3).setCellValue(temp.get(356) + milpeo);
        sheet.getRow(10).createCell(1).setCellValue(temp.get(358) + suo);
        sheet.getRow(10).createCell(2).setCellValue("");
        sheet.getRow(10).createCell(3).setCellValue(temp.get(359) + milpeo);


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
