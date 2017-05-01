import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;

/**
 * Created by sdlds on 2017/4/19.
 */


/**
 * \d+\.\d+|\d+
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
        sheet.getRow(8).createCell(3).setCellValue("");

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
        sheet.getRow(20).createCell(3).setCellValue("");
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
        sheet.getRow(10).createCell(3).setCellValue("");
        sheet.getRow(11).createCell(2).setCellValue("");
        sheet.getRow(11).createCell(3).setCellValue("");
        sheet.getRow(12).createCell(2).setCellValue("");
        sheet.getRow(12).createCell(3).setCellValue("");

        sheet.getRow(20).createCell(2).setCellValue(temp.get(86) + milton);
        sheet.getRow(20).createCell(3).setCellValue("");


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
        sheet.getRow(12).createCell(3).setCellValue(add + (Double.parseDouble(temp.get (290))*100) + per);


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

    public static void writeToExcel2008(ArrayList<String> temp, String title){
        Workbook wb = null;
        try {
            FileInputStream fileInputStream = new FileInputStream("file/" + title + ".xlsx");
            wb = WorkbookFactory.create(fileInputStream);
        }catch (Exception e){
            e.printStackTrace();
        }
        Sheet sheet = null;
        sheet = wb.getSheet("综合");
        sheet.getRow(1).createCell(2).setCellValue(temp.get(409));
        sheet.getRow(1).createCell(3).setCellValue("");
        sheet.getRow(2).createCell(2).setCellValue(temp.get(410));
        sheet.getRow(2).createCell(3).setCellValue("");
        sheet.getRow(3).createCell(2).setCellValue(temp.get(411));
        sheet.getRow(3).createCell(3).setCellValue("");
        sheet.getRow(4).createCell(2).setCellValue(temp.get(412));
        sheet.getRow(4).createCell(3).setCellValue("");
        sheet.getRow(5).createCell(2).setCellValue(temp.get(406));
        sheet.getRow(5).createCell(3).setCellValue("");
        sheet.getRow(6).createCell(2).setCellValue(temp.get(407));
        sheet.getRow(6).createCell(3).setCellValue("");
        sheet.getRow(7).createCell(2).setCellValue("");
        sheet.getRow(7).createCell(3).setCellValue("");
        sheet.getRow(8).createCell(2).setCellValue(temp.get(417) + per);
        sheet.getRow(8).createCell(3).setCellValue(add + temp.get(418) + per);

        sheet.getRow(11).createCell(2).setCellValue(temp.get(17));
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(18) + per);
        sheet.getRow(12).createCell(2).setCellValue("");
        sheet.getRow(12).createCell(3).setCellValue("");
        sheet.getRow(13).createCell(2).setCellValue(temp.get(32));
        sheet.getRow(13).createCell(3).setCellValue(add + temp.get(33) + per);
        sheet.getRow(14).createCell(2).setCellValue(temp.get(20));
        sheet.getRow(14).createCell(3).setCellValue(mis + temp.get(21) + per);
        sheet.getRow(15).createCell(2).setCellValue(temp.get(23));
        sheet.getRow(15).createCell(3).setCellValue(add + temp.get(24) + per);
        sheet.getRow(16).createCell(2).setCellValue(temp.get(26));
        sheet.getRow(16).createCell(3).setCellValue(add + temp.get(27) + per);

        sheet.getRow(20).createCell(2).setCellValue(temp.get(40) + bilyuan);
        sheet.getRow(20).createCell(3).setCellValue(mis + temp.get(41) + per);
        sheet.getRow(21).createCell(2).setCellValue("");

        sheet.getRow(30).createCell(2).setCellValue("");
        sheet.getRow(31).createCell(2).setCellValue("");
        sheet.getRow(32).createCell(2).setCellValue("");


        sheet.getRow(40).createCell(2).setCellValue("");


        sheet = wb.getSheet("农业");
        sheet.getRow(1).createCell(2).setCellValue("");
        sheet.getRow(1).createCell(3).setCellValue("");
        sheet.getRow(2).createCell(2).setCellValue(temp.get(68));
        sheet.getRow(2).createCell(3).setCellValue(add + temp.get(69) + per);
        sheet.getRow(3).createCell(2).setCellValue(temp.get(79));
        sheet.getRow(3).createCell(3).setCellValue(mis + temp.get(80) + per);
        sheet.getRow(4).createCell(2).setCellValue(temp.get(84));
        sheet.getRow(4).createCell(3).setCellValue(add + temp.get(85) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(70) + milton);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(72) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(81) + milton);
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(83) + per);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(86) + milton);
        sheet.getRow(12).createCell(3).setCellValue(mis + temp.get(87) + per);

        sheet.getRow(20).createCell(2).setCellValue(temp.get(88) + milton);
        sheet.getRow(20).createCell(3).setCellValue(mis + temp.get(90) + per);


        sheet = wb.getSheet("工业");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(131) + bilyuan);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(121) + hu);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(133) + bilyuan);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(139) + bilyuan);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(137) + bilyuan);


        sheet = wb.getSheet("投资");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(166) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(169) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(191) + bilyuan);
        sheet.getRow(10).createCell(3).setCellValue(add+ temp.get(192) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(193) + bilyuan);
        sheet.getRow(11).createCell(3).setCellValue(add+ temp.get(194) + per);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(195) + bilyuan);
        sheet.getRow(12).createCell(3).setCellValue(add+ temp.get(196) + per);


        sheet = wb.getSheet("贸易");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(197) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(198) + per);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(214) + bilyuan);
        sheet.getRow(1).createCell(3).setCellValue(add + temp.get(215) + per);
        sheet.getRow(2).createCell(2).setCellValue(temp.get(218) + bilyuan);
        sheet.getRow(2).createCell(3).setCellValue(add + temp.get(219) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(248) + bildollar);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(249) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(250) + bildollar);
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(251) + per);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(252) + bildollar);
        sheet.getRow(12).createCell(3).setCellValue(add + temp.get (253) + per);


        sheet = wb.getSheet("交通");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(258) + km);
        sheet.getRow(1).createCell(2).setCellValue("");
        sheet.getRow(2).createCell(2).setCellValue("");
        sheet.getRow(3).createCell(2).setCellValue("");
        sheet.getRow(4).createCell(2).setCellValue("");
        sheet.getRow(5).createCell(2).setCellValue("");

        sheet.getRow(10).createCell(2).setCellValue(temp.get(270) + bilyuan);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(271) + per);
        sheet.getRow(11).createCell(2).setCellValue("");
        sheet.getRow(11).createCell(3).setCellValue("");
        sheet.getRow(12).createCell(2).setCellValue(temp.get(272) + milhu);
        sheet.getRow(12).createCell(3).setCellValue(add + temp.get(273) + per);


        sheet = wb.getSheet("金融");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(304) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(305) + per);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(308) + bilyuan);
        sheet.getRow(1).createCell(3).setCellValue(add + temp.get(309) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(315) + bilyuan);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(316) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(321) + bilyuan);
        sheet.getRow(11).createCell(3).setCellValue("");


        sheet = wb.getSheet("教育");
        sheet.getRow(0).createCell(1).setCellValue("");
        sheet.getRow(0).createCell(3).setCellValue("");
        sheet.getRow(0).createCell(5).setCellValue("");


        sheet.getRow(6).createCell(1).setCellValue(temp.get(328) + suo);
        sheet.getRow(6).createCell(2).setCellValue("");
        sheet.getRow(6).createCell(3).setCellValue(temp.get(329) + milpeo);
        sheet.getRow(7).createCell(1).setCellValue(temp.get(330) + suo);
        sheet.getRow(7).createCell(2).setCellValue("");
        sheet.getRow(7).createCell(3).setCellValue(temp.get(331) + milpeo);
        sheet.getRow(8).createCell(1).setCellValue(temp.get(332) + suo);
        sheet.getRow(8).createCell(2).setCellValue("");
        sheet.getRow(8).createCell(3).setCellValue(temp.get(333) + milpeo);
        sheet.getRow(9).createCell(1).setCellValue(temp.get(337) + suo);
        sheet.getRow(9).createCell(2).setCellValue("");
        sheet.getRow(9).createCell(3).setCellValue(temp.get(338) + milpeo);
        sheet.getRow(10).createCell(1).setCellValue(temp.get(339) + suo);
        sheet.getRow(10).createCell(2).setCellValue("");
        sheet.getRow(10).createCell(3).setCellValue(temp.get(340) + milpeo);


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

    public static void writeToExcel2009(ArrayList<String> temp, String title){
        Workbook wb = null;
        try {
            FileInputStream fileInputStream = new FileInputStream("file/" + title + ".xlsx");
            wb = WorkbookFactory.create(fileInputStream);
        }catch (Exception e){
            e.printStackTrace();
        }
        Sheet sheet = null;
        sheet = wb.getSheet("综合");
        sheet.getRow(1).createCell(2).setCellValue(temp.get(393));
        sheet.getRow(1).createCell(3).setCellValue("");
        sheet.getRow(2).createCell(2).setCellValue(temp.get(394));
        sheet.getRow(2).createCell(3).setCellValue("");
        sheet.getRow(3).createCell(2).setCellValue(temp.get(395));
        sheet.getRow(3).createCell(3).setCellValue("");
        sheet.getRow(4).createCell(2).setCellValue(temp.get(396));
        sheet.getRow(4).createCell(3).setCellValue("");
        sheet.getRow(5).createCell(2).setCellValue(temp.get(390));
        sheet.getRow(5).createCell(3).setCellValue("");
        sheet.getRow(6).createCell(2).setCellValue(temp.get(391));
        sheet.getRow(6).createCell(3).setCellValue("");
        sheet.getRow(7).createCell(2).setCellValue("");
        sheet.getRow(7).createCell(3).setCellValue("");
        sheet.getRow(8).createCell(2).setCellValue(temp.get(401) + per);
        sheet.getRow(8).createCell(3).setCellValue(add + temp.get(402) + per);

        sheet.getRow(11).createCell(2).setCellValue(temp.get(7));
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(8) + per);
        sheet.getRow(12).createCell(2).setCellValue("");
        sheet.getRow(12).createCell(3).setCellValue("");
        sheet.getRow(13).createCell(2).setCellValue(temp.get(21));
        sheet.getRow(13).createCell(3).setCellValue(add + temp.get(22) + per);
        sheet.getRow(14).createCell(2).setCellValue(temp.get(9));
        sheet.getRow(14).createCell(3).setCellValue(mis + temp.get(10) + per);
        sheet.getRow(15).createCell(2).setCellValue(temp.get(11));
        sheet.getRow(15).createCell(3).setCellValue(add + temp.get(12) + per);
        sheet.getRow(16).createCell(2).setCellValue(temp.get(13));
        sheet.getRow(16).createCell(3).setCellValue(add + temp.get(14) + per);

        sheet.getRow(20).createCell(2).setCellValue(temp.get(28) + bilyuan);
        sheet.getRow(20).createCell(3).setCellValue(mis + temp.get(29) + per);
        sheet.getRow(21).createCell(2).setCellValue("");

        sheet.getRow(30).createCell(2).setCellValue("");
        sheet.getRow(31).createCell(2).setCellValue("");
        sheet.getRow(32).createCell(2).setCellValue("");


        sheet.getRow(40).createCell(2).setCellValue("");


        sheet = wb.getSheet("农业");
        sheet.getRow(1).createCell(2).setCellValue("");
        sheet.getRow(1).createCell(3).setCellValue("");
        sheet.getRow(2).createCell(2).setCellValue(temp.get(32));
        sheet.getRow(2).createCell(3).setCellValue(add + temp.get(34) + per);
        sheet.getRow(3).createCell(2).setCellValue(temp.get(35));
        sheet.getRow(3).createCell(3).setCellValue(mis + temp.get(37) + per);
        sheet.getRow(4).createCell(2).setCellValue(temp.get(38));
        sheet.getRow(4).createCell(3).setCellValue(add + temp.get(40) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(41) + milton);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(42) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(43) + milton);
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(44) + per);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(45) + milton);
        sheet.getRow(12).createCell(3).setCellValue(add + temp.get(46) + per);

        sheet.getRow(20).createCell(2).setCellValue(temp.get(53) + milton);
        sheet.getRow(20).createCell(3).setCellValue(add + temp.get(54) + per);


        sheet = wb.getSheet("工业");
        sheet.getRow(0).createCell(2).setCellValue("");
        sheet.getRow(1).createCell(2).setCellValue(temp.get(57) + hu);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(75) + bilyuan);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(77) + bilyuan);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(80) + bilyuan);


        sheet = wb.getSheet("投资");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(102) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + Double.parseDouble(temp.get(104)) * 100 + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(118) + bilyuan);
        sheet.getRow(10).createCell(3).setCellValue(add+ Double.parseDouble(temp.get(119)) * 100 + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(120) + bilyuan);
        sheet.getRow(11).createCell(3).setCellValue(add+ Double.parseDouble(temp.get(121)) * 100 + per);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(122) + bilyuan);
        sheet.getRow(12).createCell(3).setCellValue(add+ Double.parseDouble(temp.get(123)) * 100 + per);


        sheet = wb.getSheet("贸易");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(138) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(139) + per);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(140) + bilyuan);
        sheet.getRow(1).createCell(3).setCellValue(add + temp.get(141) + per);
        sheet.getRow(2).createCell(2).setCellValue(temp.get(142) + bilyuan);
        sheet.getRow(2).createCell(3).setCellValue(add + temp.get(143) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(172) + bildollar);
        sheet.getRow(10).createCell(3).setCellValue(mis + temp.get(173) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(176) + bildollar);
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(177) + per);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(174) + bildollar);
        sheet.getRow(12).createCell(3).setCellValue(mis + temp.get (175) + per);


        sheet = wb.getSheet("交通");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(180) + km);
        sheet.getRow(1).createCell(2).setCellValue("");
        sheet.getRow(2).createCell(2).setCellValue("");
        sheet.getRow(3).createCell(2).setCellValue("");
        sheet.getRow(4).createCell(2).setCellValue("");
        sheet.getRow(5).createCell(2).setCellValue("");

        sheet.getRow(10).createCell(2).setCellValue(temp.get(197) + bilyuan);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(198) + per);
        sheet.getRow(11).createCell(2).setCellValue("");
        sheet.getRow(11).createCell(3).setCellValue("");
        sheet.getRow(12).createCell(2).setCellValue(temp.get(199) + milhu);
        sheet.getRow(12).createCell(3).setCellValue(add + temp.get(200) + per);


        sheet = wb.getSheet("金融");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(222) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(223) + per);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(226) + bilyuan);
        sheet.getRow(1).createCell(3).setCellValue(add + temp.get(227) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(235) + bilyuan);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(236) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(240) + bilyuan);
        sheet.getRow(11).createCell(3).setCellValue("");


        sheet = wb.getSheet("教育");
        sheet.getRow(0).createCell(1).setCellValue("");
        sheet.getRow(0).createCell(3).setCellValue("");
        sheet.getRow(0).createCell(5).setCellValue("");


        sheet.getRow(6).createCell(1).setCellValue(temp.get(247) + suo);
        sheet.getRow(6).createCell(2).setCellValue(temp.get(257) + milpeo);
        sheet.getRow(6).createCell(3).setCellValue(temp.get(258) + milpeo);
        sheet.getRow(7).createCell(1).setCellValue("");
        sheet.getRow(7).createCell(2).setCellValue("");
        sheet.getRow(7).createCell(3).setCellValue("");
        sheet.getRow(8).createCell(1).setCellValue("");
        sheet.getRow(8).createCell(2).setCellValue("");
        sheet.getRow(8).createCell(3).setCellValue("");
        sheet.getRow(9).createCell(1).setCellValue(temp.get(245) + suo);
        sheet.getRow(9).createCell(2).setCellValue(temp.get(251) + milpeo);
        sheet.getRow(9).createCell(3).setCellValue(temp.get(252) + milpeo);
        sheet.getRow(10).createCell(1).setCellValue(temp.get(244) + suo);
        sheet.getRow(10).createCell(2).setCellValue("1.6" + milpeo);
        sheet.getRow(10).createCell(3).setCellValue(temp.get(249) + milpeo);


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

    public static void writeToExcel2010(ArrayList<String> temp, String title){
        Workbook wb = null;
        try {
            FileInputStream fileInputStream = new FileInputStream("file/" + title + ".xlsx");
            wb = WorkbookFactory.create(fileInputStream);
        }catch (Exception e){
            e.printStackTrace();
        }
        Sheet sheet = null;
        sheet = wb.getSheet("综合");
        sheet.getRow(1).createCell(2).setCellValue("");
        sheet.getRow(1).createCell(3).setCellValue("");
        sheet.getRow(2).createCell(2).setCellValue("");
        sheet.getRow(2).createCell(3).setCellValue("");
        sheet.getRow(3).createCell(2).setCellValue("");
        sheet.getRow(3).createCell(3).setCellValue("");
        sheet.getRow(4).createCell(2).setCellValue("");
        sheet.getRow(4).createCell(3).setCellValue("");
        sheet.getRow(5).createCell(2).setCellValue("");
        sheet.getRow(5).createCell(3).setCellValue("");
        sheet.getRow(6).createCell(2).setCellValue("");
        sheet.getRow(6).createCell(3).setCellValue("");
        sheet.getRow(7).createCell(2).setCellValue("");
        sheet.getRow(7).createCell(3).setCellValue("");
        sheet.getRow(8).createCell(2).setCellValue("");
        sheet.getRow(8).createCell(3).setCellValue("");

        sheet.getRow(11).createCell(2).setCellValue(temp.get(7));
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(8) + per);
        sheet.getRow(12).createCell(2).setCellValue("");
        sheet.getRow(12).createCell(3).setCellValue("");
        sheet.getRow(13).createCell(2).setCellValue("");
        sheet.getRow(13).createCell(3).setCellValue("");
        sheet.getRow(14).createCell(2).setCellValue(temp.get(9));
        sheet.getRow(14).createCell(3).setCellValue(add + temp.get(10) + per);
        sheet.getRow(15).createCell(2).setCellValue(temp.get(11));
        sheet.getRow(15).createCell(3).setCellValue(add + temp.get(12) + per);
        sheet.getRow(16).createCell(2).setCellValue(temp.get(13));
        sheet.getRow(16).createCell(3).setCellValue(add + temp.get(14) + per);

        sheet.getRow(20).createCell(2).setCellValue(temp.get(21) + bilyuan);
        sheet.getRow(20).createCell(3).setCellValue(add + temp.get(22) + per);
        sheet.getRow(21).createCell(2).setCellValue("");

        sheet.getRow(30).createCell(2).setCellValue("");
        sheet.getRow(31).createCell(2).setCellValue("");
        sheet.getRow(32).createCell(2).setCellValue("");


        sheet.getRow(40).createCell(2).setCellValue("");


        sheet = wb.getSheet("农业");
        sheet.getRow(1).createCell(2).setCellValue("");
        sheet.getRow(1).createCell(3).setCellValue("");
        sheet.getRow(2).createCell(2).setCellValue(temp.get(25));
        sheet.getRow(2).createCell(3).setCellValue(add + temp.get(27) + per);
        sheet.getRow(3).createCell(2).setCellValue(temp.get(28));
        sheet.getRow(3).createCell(3).setCellValue(add + temp.get(30) + per);
        sheet.getRow(4).createCell(2).setCellValue(temp.get(31));
        sheet.getRow(4).createCell(3).setCellValue(mis + temp.get(33) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(34) + milton);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(36) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(41) + milton);
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(42) + per);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(43) + milton);
        sheet.getRow(12).createCell(3).setCellValue(mis + temp.get(44) + per);

        sheet.getRow(20).createCell(2).setCellValue(temp.get(51) + milton);
        sheet.getRow(20).createCell(3).setCellValue(add + temp.get(52) + per);


        sheet = wb.getSheet("工业");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(55) + bilyuan);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(59) + hu);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(75) + bilyuan);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(77) + bilyuan);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(79) + bilyuan);


        sheet = wb.getSheet("投资");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(103) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue("");

        sheet.getRow(10).createCell(2).setCellValue("");
        sheet.getRow(10).createCell(3).setCellValue("");
        sheet.getRow(11).createCell(2).setCellValue("");
        sheet.getRow(11).createCell(3).setCellValue("");
        sheet.getRow(12).createCell(2).setCellValue("");
        sheet.getRow(12).createCell(3).setCellValue("");


        sheet = wb.getSheet("贸易");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(123) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(124) + per);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(125) + bilyuan);
        sheet.getRow(1).createCell(3).setCellValue(add + temp.get(126) + per);
        sheet.getRow(2).createCell(2).setCellValue(temp.get(127) + bilyuan);
        sheet.getRow(2).createCell(3).setCellValue(add + temp.get(128) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(155) + bildollar);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(156) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(159) + bildollar);
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(160) + per);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(157) + bildollar);
        sheet.getRow(12).createCell(3).setCellValue(add + temp.get (158) + per);


        sheet = wb.getSheet("交通");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(163) + km);
        sheet.getRow(1).createCell(2).setCellValue("");
        sheet.getRow(2).createCell(2).setCellValue("");
        sheet.getRow(3).createCell(2).setCellValue("");
        sheet.getRow(4).createCell(2).setCellValue("");
        sheet.getRow(5).createCell(2).setCellValue("");

        sheet.getRow(10).createCell(2).setCellValue("");
        sheet.getRow(10).createCell(3).setCellValue("");
        sheet.getRow(11).createCell(2).setCellValue("");
        sheet.getRow(11).createCell(3).setCellValue("");
        sheet.getRow(12).createCell(2).setCellValue(temp.get(178) + milhu);
        sheet.getRow(12).createCell(3).setCellValue(add + temp.get(179) + per);


        sheet = wb.getSheet("金融");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(186) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(187) + per);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(190) + bilyuan);
        sheet.getRow(1).createCell(3).setCellValue(add + temp.get(191) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(201) + bilyuan);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(202) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(205) + bilyuan);
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(206) + per);


        sheet = wb.getSheet("教育");
        sheet.getRow(0).createCell(1).setCellValue("");
        sheet.getRow(0).createCell(3).setCellValue("");
        sheet.getRow(0).createCell(5).setCellValue("");


        sheet.getRow(6).createCell(1).setCellValue(temp.get(212) + suo);
        sheet.getRow(6).createCell(2).setCellValue(temp.get(220) + milpeo);
        sheet.getRow(6).createCell(3).setCellValue(temp.get(221) + milpeo);
        sheet.getRow(7).createCell(1).setCellValue("");
        sheet.getRow(7).createCell(2).setCellValue("");
        sheet.getRow(7).createCell(3).setCellValue("");
        sheet.getRow(8).createCell(1).setCellValue("");
        sheet.getRow(8).createCell(2).setCellValue("");
        sheet.getRow(8).createCell(3).setCellValue("");
        sheet.getRow(9).createCell(1).setCellValue(temp.get(210) + suo);
        sheet.getRow(9).createCell(2).setCellValue(temp.get(216) + milpeo);
        sheet.getRow(9).createCell(3).setCellValue(temp.get(217) + milpeo);
        sheet.getRow(10).createCell(1).setCellValue(temp.get(209) + suo);
        sheet.getRow(10).createCell(2).setCellValue(temp.get(213) + milpeo);
        sheet.getRow(10).createCell(3).setCellValue(temp.get(214) + milpeo);


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

    public static void writeToExcel2011(ArrayList<String> temp, String title){
        Workbook wb = null;
        try {
            FileInputStream fileInputStream = new FileInputStream("file/" + title + ".xlsx");
            wb = WorkbookFactory.create(fileInputStream);
        }catch (Exception e){
            e.printStackTrace();
        }
        Sheet sheet = null;
        sheet = wb.getSheet("综合");
        sheet.getRow(1).createCell(2).setCellValue(temp.get(344));
        sheet.getRow(1).createCell(3).setCellValue("");
        sheet.getRow(2).createCell(2).setCellValue(temp.get(345));
        sheet.getRow(2).createCell(3).setCellValue("");
        sheet.getRow(3).createCell(2).setCellValue("");
        sheet.getRow(3).createCell(3).setCellValue("");
        sheet.getRow(4).createCell(2).setCellValue("");
        sheet.getRow(4).createCell(3).setCellValue("");
        sheet.getRow(5).createCell(2).setCellValue(temp.get(342));
        sheet.getRow(5).createCell(3).setCellValue("");
        sheet.getRow(6).createCell(2).setCellValue(temp.get(343));
        sheet.getRow(6).createCell(3).setCellValue("");
        sheet.getRow(7).createCell(2).setCellValue("");
        sheet.getRow(7).createCell(3).setCellValue("");
        sheet.getRow(8).createCell(2).setCellValue(temp.get(350) + per);
        sheet.getRow(8).createCell(3).setCellValue(add + temp.get(351) + per);

        sheet.getRow(11).createCell(2).setCellValue(temp.get(3));
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(4) + per);
        sheet.getRow(12).createCell(2).setCellValue("");
        sheet.getRow(12).createCell(3).setCellValue("");
        sheet.getRow(13).createCell(2).setCellValue(temp.get(14));
        sheet.getRow(13).createCell(3).setCellValue("");
        sheet.getRow(14).createCell(2).setCellValue("");
        sheet.getRow(14).createCell(3).setCellValue(add + temp.get(6) + per);
        sheet.getRow(15).createCell(2).setCellValue("");
        sheet.getRow(15).createCell(3).setCellValue(add + temp.get(7) + per);
        sheet.getRow(16).createCell(2).setCellValue("");
        sheet.getRow(16).createCell(3).setCellValue(add + temp.get(8) + per);

        sheet.getRow(20).createCell(2).setCellValue(temp.get(19) + bilyuan);
        sheet.getRow(20).createCell(3).setCellValue(mis + temp.get(20) + per);
        sheet.getRow(21).createCell(2).setCellValue("");

        sheet.getRow(30).createCell(2).setCellValue("");
        sheet.getRow(31).createCell(2).setCellValue("");
        sheet.getRow(32).createCell(2).setCellValue("");


        sheet.getRow(40).createCell(2).setCellValue("");


        sheet = wb.getSheet("农业");
        sheet.getRow(1).createCell(2).setCellValue("");
        sheet.getRow(1).createCell(3).setCellValue("");
        sheet.getRow(2).createCell(2).setCellValue(temp.get(41));
        sheet.getRow(2).createCell(3).setCellValue(mis + temp.get(43) + per);
        sheet.getRow(3).createCell(2).setCellValue(temp.get(44));
        sheet.getRow(3).createCell(3).setCellValue(mis + temp.get(46) + per);
        sheet.getRow(4).createCell(2).setCellValue(temp.get(47));
        sheet.getRow(4).createCell(3).setCellValue(add + temp.get(49) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(50) + milton);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(52) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(53) + milton);
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(54) + per);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(55) + milton);
        sheet.getRow(12).createCell(3).setCellValue(add + temp.get(56) + per);

        sheet.getRow(20).createCell(2).setCellValue(temp.get(63) + milton);
        sheet.getRow(20).createCell(3).setCellValue(mis + temp.get(64) + per);


        sheet = wb.getSheet("工业");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(72) + bilyuan);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(75) + hu);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(81) + bilyuan);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(83) + bilyuan);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(85) + bilyuan);


        sheet = wb.getSheet("投资");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(117) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(118) + per);

        sheet.getRow(10).createCell(2).setCellValue("");
        sheet.getRow(10).createCell(3).setCellValue("");
        sheet.getRow(11).createCell(2).setCellValue("");
        sheet.getRow(11).createCell(3).setCellValue("");
        sheet.getRow(12).createCell(2).setCellValue("");
        sheet.getRow(12).createCell(3).setCellValue("");


        sheet = wb.getSheet("贸易");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(133) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(134) + per);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(135) + bilyuan);
        sheet.getRow(1).createCell(3).setCellValue(add + temp.get(136) + per);
        sheet.getRow(2).createCell(2).setCellValue(temp.get(137) + bilyuan);
        sheet.getRow(2).createCell(3).setCellValue(add + temp.get(138) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(148) + bildollar);
        sheet.getRow(10).createCell(3).setCellValue(mis + temp.get(149) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(150) + bildollar);
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(151) + per);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(152) + bildollar);
        sheet.getRow(12).createCell(3).setCellValue(mis + temp.get (153) + per);


        sheet = wb.getSheet("交通");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(161) + km);
        sheet.getRow(1).createCell(2).setCellValue("");
        sheet.getRow(2).createCell(2).setCellValue(temp.get(169) + "万吨");
        sheet.getRow(3).createCell(2).setCellValue("");
        sheet.getRow(4).createCell(2).setCellValue("");
        sheet.getRow(5).createCell(2).setCellValue("");

        sheet.getRow(10).createCell(2).setCellValue(temp.get(179) + bilyuan);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(180) + per);
        sheet.getRow(11).createCell(2).setCellValue("");
        sheet.getRow(11).createCell(3).setCellValue("");
        sheet.getRow(12).createCell(2).setCellValue(temp.get(183) + milhu);
        sheet.getRow(12).createCell(3).setCellValue(add + temp.get(184) + per);


        sheet = wb.getSheet("金融");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(197) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(198) + per);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(201) + bilyuan);
        sheet.getRow(1).createCell(3).setCellValue(add + temp.get(202) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(210) + bilyuan);
        sheet.getRow(10).createCell(3).setCellValue(mis + temp.get(211) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(216) + bilyuan);
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(217) + per);


        sheet = wb.getSheet("教育");
        sheet.getRow(0).createCell(1).setCellValue("");
        sheet.getRow(0).createCell(3).setCellValue("");
        sheet.getRow(0).createCell(5).setCellValue("");


        sheet.getRow(6).createCell(1).setCellValue(temp.get(227) + suo);
        sheet.getRow(6).createCell(2).setCellValue(temp.get(240) + milpeo);
        sheet.getRow(6).createCell(3).setCellValue(temp.get(241) + milpeo);
        sheet.getRow(7).createCell(1).setCellValue("");
        sheet.getRow(7).createCell(2).setCellValue("");
        sheet.getRow(7).createCell(3).setCellValue("");
        sheet.getRow(8).createCell(1).setCellValue("");
        sheet.getRow(8).createCell(2).setCellValue("");
        sheet.getRow(8).createCell(3).setCellValue("");
        sheet.getRow(9).createCell(1).setCellValue(temp.get(225) + suo);
        sheet.getRow(9).createCell(2).setCellValue(temp.get(232) + milpeo);
        sheet.getRow(9).createCell(3).setCellValue(temp.get(233) + milpeo);
        sheet.getRow(10).createCell(1).setCellValue(temp.get(224) + suo);
        sheet.getRow(10).createCell(2).setCellValue(temp.get(228) + milpeo);
        sheet.getRow(10).createCell(3).setCellValue(temp.get(229) + milpeo);


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

    public static void writeToExcel2012(ArrayList<String> temp, String title){
        Workbook wb = null;
        try {
            FileInputStream fileInputStream = new FileInputStream("file/" + title + ".xlsx");
            wb = WorkbookFactory.create(fileInputStream);
        }catch (Exception e){
            e.printStackTrace();
        }
        Sheet sheet = null;
        sheet = wb.getSheet("综合");
        sheet.getRow(1).createCell(2).setCellValue(temp.get(339));
        sheet.getRow(1).createCell(3).setCellValue("");
        sheet.getRow(2).createCell(2).setCellValue(temp.get(340));
        sheet.getRow(2).createCell(3).setCellValue("");
        sheet.getRow(3).createCell(2).setCellValue("");
        sheet.getRow(3).createCell(3).setCellValue("");
        sheet.getRow(4).createCell(2).setCellValue("");
        sheet.getRow(4).createCell(3).setCellValue("");
        sheet.getRow(5).createCell(2).setCellValue(temp.get(337));
        sheet.getRow(5).createCell(3).setCellValue("");
        sheet.getRow(6).createCell(2).setCellValue(temp.get(338));
        sheet.getRow(6).createCell(3).setCellValue("");
        sheet.getRow(7).createCell(2).setCellValue("");
        sheet.getRow(7).createCell(3).setCellValue("");
        sheet.getRow(8).createCell(2).setCellValue(temp.get(345) + per);
        sheet.getRow(8).createCell(3).setCellValue(add + temp.get(346) + per);

        sheet.getRow(11).createCell(2).setCellValue(temp.get(7));
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(8) + per);
        sheet.getRow(12).createCell(2).setCellValue("");
        sheet.getRow(12).createCell(3).setCellValue("");
        sheet.getRow(13).createCell(2).setCellValue(temp.get(14));
        sheet.getRow(13).createCell(3).setCellValue("");
        sheet.getRow(14).createCell(2).setCellValue("");
        sheet.getRow(14).createCell(3).setCellValue(add + temp.get(9) + per);
        sheet.getRow(15).createCell(2).setCellValue("");
        sheet.getRow(15).createCell(3).setCellValue(add + temp.get(10) + per);
        sheet.getRow(16).createCell(2).setCellValue("");
        sheet.getRow(16).createCell(3).setCellValue(add + temp.get(11) + per);

        sheet.getRow(20).createCell(2).setCellValue(temp.get(15) + bilyuan);
        sheet.getRow(20).createCell(3).setCellValue(add + temp.get(16) + per);
        sheet.getRow(21).createCell(2).setCellValue("");

        sheet.getRow(30).createCell(2).setCellValue("");
        sheet.getRow(31).createCell(2).setCellValue("");
        sheet.getRow(32).createCell(2).setCellValue("");


        sheet.getRow(40).createCell(2).setCellValue("");


        sheet = wb.getSheet("农业");
        sheet.getRow(1).createCell(2).setCellValue(temp.get(20));
        sheet.getRow(1).createCell(3).setCellValue("");
        sheet.getRow(2).createCell(2).setCellValue(temp.get(23));
        sheet.getRow(2).createCell(3).setCellValue("");
        sheet.getRow(3).createCell(2).setCellValue(temp.get(26));
        sheet.getRow(3).createCell(3).setCellValue("");
        sheet.getRow(4).createCell(2).setCellValue(temp.get(29));
        sheet.getRow(4).createCell(3).setCellValue("");

        sheet.getRow(10).createCell(2).setCellValue(temp.get(32) + milton);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(35) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(36) + milton);
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(39) + per);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(40) + milton);
        sheet.getRow(12).createCell(3).setCellValue(add + temp.get(43) + per);

        sheet.getRow(20).createCell(2).setCellValue("");
        sheet.getRow(20).createCell(3).setCellValue(add + temp.get(45) + per);


        sheet = wb.getSheet("工业");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(48) + bilyuan);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(50) + hu);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(66) + bilyuan);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(68) + bilyuan);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(70) + bilyuan);


        sheet = wb.getSheet("投资");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(90) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(91) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(107) + bilyuan);
        sheet.getRow(10).createCell(3).setCellValue(mis+ temp.get(108) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(109) + bilyuan);
        sheet.getRow(11).createCell(3).setCellValue(add+ temp.get(110) + per);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(111) + bilyuan);
        sheet.getRow(12).createCell(3).setCellValue(add+ temp.get(112) + per);


        sheet = wb.getSheet("贸易");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(131) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(132) + per);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(133) + bilyuan);
        sheet.getRow(1).createCell(3).setCellValue(add + temp.get(134) + per);
        sheet.getRow(2).createCell(2).setCellValue(temp.get(135) + bilyuan);
        sheet.getRow(2).createCell(3).setCellValue(add + temp.get(136) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(141) + bildollar);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(142) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(151) + bildollar);
        sheet.getRow(11).createCell(3).setCellValue(mis + temp.get(152) + per);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(143) + bildollar);
        sheet.getRow(12).createCell(3).setCellValue(add + temp.get (144) + per);


        sheet = wb.getSheet("交通");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(155) + km);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(159) + km);
        sheet.getRow(2).createCell(2).setCellValue(temp.get(163) + "万吨");
        sheet.getRow(3).createCell(2).setCellValue("");
        sheet.getRow(4).createCell(2).setCellValue("");
        sheet.getRow(5).createCell(2).setCellValue("");

        sheet.getRow(10).createCell(2).setCellValue(temp.get(171) + bilyuan);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(172) + per);
        sheet.getRow(11).createCell(2).setCellValue("");
        sheet.getRow(11).createCell(3).setCellValue("");
        sheet.getRow(12).createCell(2).setCellValue(temp.get(175) + milhu);
        sheet.getRow(12).createCell(3).setCellValue(add + temp.get(176) + per);


        sheet = wb.getSheet("金融");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(207) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(208) + per);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(211) + bilyuan);
        sheet.getRow(1).createCell(3).setCellValue(add + temp.get(212) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(220) + bilyuan);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(221) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(226) + bilyuan);
        sheet.getRow(11).createCell(3).setCellValue(mis + temp.get(227) + per);


        sheet = wb.getSheet("教育");
        sheet.getRow(0).createCell(1).setCellValue("");
        sheet.getRow(0).createCell(3).setCellValue("");
        sheet.getRow(0).createCell(5).setCellValue("");


        sheet.getRow(6).createCell(1).setCellValue(temp.get(237) + suo);
        sheet.getRow(6).createCell(2).setCellValue(temp.get(250) + milpeo);
        sheet.getRow(6).createCell(3).setCellValue(temp.get(251) + milpeo);
        sheet.getRow(7).createCell(1).setCellValue("");
        sheet.getRow(7).createCell(2).setCellValue("");
        sheet.getRow(7).createCell(3).setCellValue("");
        sheet.getRow(8).createCell(1).setCellValue("");
        sheet.getRow(8).createCell(2).setCellValue("");
        sheet.getRow(8).createCell(3).setCellValue("");
        sheet.getRow(9).createCell(1).setCellValue(temp.get(235) + suo);
        sheet.getRow(9).createCell(2).setCellValue(temp.get(242) + milpeo);
        sheet.getRow(9).createCell(3).setCellValue(temp.get(243) + milpeo);
        sheet.getRow(10).createCell(1).setCellValue(temp.get(234) + suo);
        sheet.getRow(10).createCell(2).setCellValue(temp.get(238) + milpeo);
        sheet.getRow(10).createCell(3).setCellValue(temp.get(239) + milpeo);


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

    public static void writeToExcel2013(ArrayList<String> temp, String title){
        Workbook wb = null;
        try {
            FileInputStream fileInputStream = new FileInputStream("file/" + title + ".xlsx");
            wb = WorkbookFactory.create(fileInputStream);
        }catch (Exception e){
            e.printStackTrace();
        }
        Sheet sheet = null;
        sheet = wb.getSheet("综合");
        sheet.getRow(1).createCell(2).setCellValue(temp.get(3));
        sheet.getRow(1).createCell(3).setCellValue("");
        sheet.getRow(2).createCell(2).setCellValue(temp.get(4));
        sheet.getRow(2).createCell(3).setCellValue("");
        sheet.getRow(3).createCell(2).setCellValue("");
        sheet.getRow(3).createCell(3).setCellValue("");
        sheet.getRow(4).createCell(2).setCellValue("");
        sheet.getRow(4).createCell(3).setCellValue("");
        sheet.getRow(5).createCell(2).setCellValue(temp.get(5));
        sheet.getRow(5).createCell(3).setCellValue("");
        sheet.getRow(6).createCell(2).setCellValue(temp.get(6));
        sheet.getRow(6).createCell(3).setCellValue("");
        sheet.getRow(7).createCell(2).setCellValue(temp.get(7));
        sheet.getRow(7).createCell(3).setCellValue("");
        sheet.getRow(8).createCell(2).setCellValue(temp.get(9) + per);
        sheet.getRow(8).createCell(3).setCellValue("");

        sheet.getRow(11).createCell(2).setCellValue(temp.get(10));
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(11) + per);
        sheet.getRow(12).createCell(2).setCellValue("");
        sheet.getRow(12).createCell(3).setCellValue("");
        sheet.getRow(13).createCell(2).setCellValue(temp.get(21));
        sheet.getRow(13).createCell(3).setCellValue("");
        sheet.getRow(14).createCell(2).setCellValue(temp.get(12));
        sheet.getRow(14).createCell(3).setCellValue(add + temp.get(13) + per);
        sheet.getRow(15).createCell(2).setCellValue(temp.get(14));
        sheet.getRow(15).createCell(3).setCellValue(add + temp.get(15) + per);
        sheet.getRow(16).createCell(2).setCellValue(temp.get(16));
        sheet.getRow(16).createCell(3).setCellValue(add + temp.get(17) + per);

        sheet.getRow(20).createCell(2).setCellValue(temp.get(22) + bilyuan);
        sheet.getRow(20).createCell(3).setCellValue(add + temp.get(23) + per);
        sheet.getRow(21).createCell(2).setCellValue("");

        sheet.getRow(30).createCell(2).setCellValue("");
        sheet.getRow(31).createCell(2).setCellValue("");
        sheet.getRow(32).createCell(2).setCellValue("");


        sheet.getRow(40).createCell(2).setCellValue("");


        sheet = wb.getSheet("农业");
        sheet.getRow(1).createCell(2).setCellValue(temp.get(32));
        sheet.getRow(1).createCell(3).setCellValue("");
        sheet.getRow(2).createCell(2).setCellValue(temp.get(34));
        sheet.getRow(2).createCell(3).setCellValue("");
        sheet.getRow(3).createCell(2).setCellValue(temp.get(36));
        sheet.getRow(3).createCell(3).setCellValue("");
        sheet.getRow(4).createCell(2).setCellValue(temp.get(38));
        sheet.getRow(4).createCell(3).setCellValue("");

        sheet.getRow(10).createCell(2).setCellValue(temp.get(40) + milton);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(42) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(43) + milton);
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(44) + per);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(45) + milton);
        sheet.getRow(12).createCell(3).setCellValue(add + temp.get(46) + per);

        sheet.getRow(20).createCell(2).setCellValue(temp.get(47) + milton);
        sheet.getRow(20).createCell(3).setCellValue(add + temp.get(48) + per);


        sheet = wb.getSheet("工业");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(72) + bilyuan);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(53) + hu);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(91) + bilyuan);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(93) + bilyuan);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(95) + bilyuan);


        sheet = wb.getSheet("投资");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(101) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(102) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(109) + bilyuan);
        sheet.getRow(10).createCell(3).setCellValue(mis+ temp.get(110) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(111) + bilyuan);
        sheet.getRow(11).createCell(3).setCellValue(add+ temp.get(112) + per);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(115) + bilyuan);
        sheet.getRow(12).createCell(3).setCellValue(mis+ temp.get(116) + per);


        sheet = wb.getSheet("贸易");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(159) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(160) + per);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(161) + bilyuan);
        sheet.getRow(1).createCell(3).setCellValue(add + temp.get(162) + per);
        sheet.getRow(2).createCell(2).setCellValue(temp.get(163) + bilyuan);
        sheet.getRow(2).createCell(3).setCellValue(add + temp.get(164) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(171) + bildollar);
        sheet.getRow(10).createCell(3).setCellValue(mis + temp.get(172) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(178) + bildollar);
        sheet.getRow(11).createCell(3).setCellValue(mis + temp.get(179) + per);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(173) + bildollar);
        sheet.getRow(12).createCell(3).setCellValue(add + temp.get (174) + per);


        sheet = wb.getSheet("交通");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(155) + km);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(159) + km);
        sheet.getRow(2).createCell(2).setCellValue(temp.get(164) + "万吨");
        sheet.getRow(3).createCell(2).setCellValue(temp.get(166) + "亿人");
        sheet.getRow(4).createCell(2).setCellValue(temp.get(168) + "万人");
        sheet.getRow(5).createCell(2).setCellValue(temp.get(169) + "万吨");

        sheet.getRow(10).createCell(2).setCellValue("");
        sheet.getRow(10).createCell(3).setCellValue("");
        sheet.getRow(11).createCell(2).setCellValue("");
        sheet.getRow(11).createCell(3).setCellValue("");
        sheet.getRow(12).createCell(2).setCellValue(temp.get(170) + milhu);
        sheet.getRow(12).createCell(3).setCellValue("");


        sheet = wb.getSheet("金融");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(182) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(183) + per);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(184) + bilyuan);
        sheet.getRow(1).createCell(3).setCellValue(add + temp.get(186) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(203) + bilyuan);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(204) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(207) + bilyuan);
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(208) + per);


        sheet = wb.getSheet("教育");
        sheet.getRow(0).createCell(1).setCellValue("665" + suo);
        sheet.getRow(0).createCell(3).setCellValue("50.8" + milpeo);
        sheet.getRow(0).createCell(5).setCellValue("2.9" + milpeo);


        sheet.getRow(6).createCell(1).setCellValue(temp.get(225) + suo);
        sheet.getRow(6).createCell(2).setCellValue("");
        sheet.getRow(6).createCell(3).setCellValue("15.9" + milpeo);
        sheet.getRow(7).createCell(1).setCellValue("");
        sheet.getRow(7).createCell(2).setCellValue("");
        sheet.getRow(7).createCell(3).setCellValue("");
        sheet.getRow(8).createCell(1).setCellValue("");
        sheet.getRow(8).createCell(2).setCellValue("");
        sheet.getRow(8).createCell(3).setCellValue("");
        sheet.getRow(9).createCell(1).setCellValue(temp.get(227) + suo);
        sheet.getRow(9).createCell(2).setCellValue("");
        sheet.getRow(9).createCell(3).setCellValue("4.9" + milpeo);
        sheet.getRow(10).createCell(1).setCellValue(temp.get(228) + suo);
        sheet.getRow(10).createCell(2).setCellValue("");
        sheet.getRow(10).createCell(3).setCellValue("6.2" + milpeo);


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

    public static void writeToExcel2014(ArrayList<String> temp, String title){
        Workbook wb = null;
        try {
            FileInputStream fileInputStream = new FileInputStream("file/" + title + ".xlsx");
            wb = WorkbookFactory.create(fileInputStream);
        }catch (Exception e){
            e.printStackTrace();
        }
        Sheet sheet = null;
        sheet = wb.getSheet("综合");
        sheet.getRow(1).createCell(2).setCellValue(temp.get(6));
        sheet.getRow(1).createCell(3).setCellValue("");
        sheet.getRow(2).createCell(2).setCellValue(temp.get(7));
        sheet.getRow(2).createCell(3).setCellValue("");
        sheet.getRow(3).createCell(2).setCellValue("");
        sheet.getRow(3).createCell(3).setCellValue("");
        sheet.getRow(4).createCell(2).setCellValue("");
        sheet.getRow(4).createCell(3).setCellValue("");
        sheet.getRow(5).createCell(2).setCellValue(temp.get(8));
        sheet.getRow(5).createCell(3).setCellValue("");
        sheet.getRow(6).createCell(2).setCellValue(temp.get(9));
        sheet.getRow(6).createCell(3).setCellValue("");
        sheet.getRow(7).createCell(2).setCellValue(temp.get(10));
        sheet.getRow(7).createCell(3).setCellValue("");
        sheet.getRow(8).createCell(2).setCellValue(temp.get(12) + per);
        sheet.getRow(8).createCell(3).setCellValue("");

        sheet.getRow(11).createCell(2).setCellValue(temp.get(13));
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(14) + per);
        sheet.getRow(12).createCell(2).setCellValue("");
        sheet.getRow(12).createCell(3).setCellValue("");
        sheet.getRow(13).createCell(2).setCellValue("");
        sheet.getRow(13).createCell(3).setCellValue("");
        sheet.getRow(14).createCell(2).setCellValue(temp.get(15));
        sheet.getRow(14).createCell(3).setCellValue(add + temp.get(16) + per);
        sheet.getRow(15).createCell(2).setCellValue(temp.get(17));
        sheet.getRow(15).createCell(3).setCellValue(add + temp.get(18) + per);
        sheet.getRow(16).createCell(2).setCellValue(temp.get(19));
        sheet.getRow(16).createCell(3).setCellValue(add + temp.get(20) + per);

        sheet.getRow(20).createCell(2).setCellValue(temp.get(24) + bilyuan);
        sheet.getRow(20).createCell(3).setCellValue(add + temp.get(25) + per);
        sheet.getRow(21).createCell(2).setCellValue("");

        sheet.getRow(30).createCell(2).setCellValue("");
        sheet.getRow(31).createCell(2).setCellValue("");
        sheet.getRow(32).createCell(2).setCellValue("");


        sheet.getRow(40).createCell(2).setCellValue("");


        sheet = wb.getSheet("农业");
        sheet.getRow(1).createCell(2).setCellValue(temp.get(39));
        sheet.getRow(1).createCell(3).setCellValue(add + temp.get(41) + per);
        sheet.getRow(2).createCell(2).setCellValue(temp.get(42));
        sheet.getRow(2).createCell(3).setCellValue(add + temp.get(44) + per);
        sheet.getRow(3).createCell(2).setCellValue(temp.get(45));
        sheet.getRow(3).createCell(3).setCellValue(mis + temp.get(47) + per);
        sheet.getRow(4).createCell(2).setCellValue(temp.get(48));
        sheet.getRow(4).createCell(3).setCellValue(add + temp.get(50) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(51) + milton);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(53) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(54) + milton);
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(56) + per);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(57) + milton);
        sheet.getRow(12).createCell(3).setCellValue(add + temp.get(59) + per);

        sheet.getRow(20).createCell(2).setCellValue(temp.get(60) + milton);
        sheet.getRow(20).createCell(3).setCellValue(add + temp.get(62) + per);


        sheet = wb.getSheet("工业");
        sheet.getRow(0).createCell(2).setCellValue("");
        sheet.getRow(1).createCell(2).setCellValue(temp.get(69) + hu);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(79) + bilyuan);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(81) + bilyuan);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(83) + bilyuan);


        sheet = wb.getSheet("投资");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(118) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(119) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(122) + bilyuan);
        sheet.getRow(10).createCell(3).setCellValue(mis+ temp.get(123) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(124) + bilyuan);
        sheet.getRow(11).createCell(3).setCellValue(add+ temp.get(125) + per);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(128) + bilyuan);
        sheet.getRow(12).createCell(3).setCellValue(add+ temp.get(129) + per);


        sheet = wb.getSheet("贸易");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(140) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(141) + per);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(142) + bilyuan);
        sheet.getRow(1).createCell(3).setCellValue(add + temp.get(143) + per);
        sheet.getRow(2).createCell(2).setCellValue(temp.get(144) + bilyuan);
        sheet.getRow(2).createCell(3).setCellValue(add + temp.get(145) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(150) + bildollar);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(151) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(157) + bildollar);
        sheet.getRow(11).createCell(3).setCellValue(mis + temp.get(158) + per);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(152) + bildollar);
        sheet.getRow(12).createCell(3).setCellValue(add + temp.get (153) + per);


        sheet = wb.getSheet("交通");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(161) + km);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(166) + km);
        sheet.getRow(2).createCell(2).setCellValue("");
        sheet.getRow(3).createCell(2).setCellValue("");
        sheet.getRow(4).createCell(2).setCellValue(temp.get(177) + "万人");
        sheet.getRow(5).createCell(2).setCellValue(temp.get(178) + "万吨");

        sheet.getRow(10).createCell(2).setCellValue(temp.get(170) + bilyuan);
        sheet.getRow(10).createCell(3).setCellValue("");
        sheet.getRow(11).createCell(2).setCellValue(temp.get(171) + milhu);
        sheet.getRow(11).createCell(3).setCellValue("");
        sheet.getRow(12).createCell(2).setCellValue(temp.get(172) + milhu);
        sheet.getRow(12).createCell(3).setCellValue("");


        sheet = wb.getSheet("金融");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(185) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(186) + per);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(189) + bilyuan);
        sheet.getRow(1).createCell(3).setCellValue(add + temp.get(190) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(210) + bilyuan);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(211) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(214) + bilyuan);
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(215) + per);


        sheet = wb.getSheet("教育");
        sheet.getRow(0).createCell(1).setCellValue("777" + suo);
        sheet.getRow(0).createCell(3).setCellValue("49.9" + milpeo);
        sheet.getRow(0).createCell(5).setCellValue("2.9" + milpeo);


        sheet.getRow(6).createCell(1).setCellValue("");
        sheet.getRow(6).createCell(2).setCellValue("");
        sheet.getRow(6).createCell(3).setCellValue("16.2" + milpeo);
        sheet.getRow(7).createCell(1).setCellValue("");
        sheet.getRow(7).createCell(2).setCellValue("");
        sheet.getRow(7).createCell(3).setCellValue("");
        sheet.getRow(8).createCell(1).setCellValue("");
        sheet.getRow(8).createCell(2).setCellValue("");
        sheet.getRow(8).createCell(3).setCellValue("");
        sheet.getRow(9).createCell(1).setCellValue("");
        sheet.getRow(9).createCell(2).setCellValue("");
        sheet.getRow(9).createCell(3).setCellValue("3.6" + milpeo);
        sheet.getRow(10).createCell(1).setCellValue("");
        sheet.getRow(10).createCell(2).setCellValue("");
        sheet.getRow(10).createCell(3).setCellValue("6.7" + milpeo);


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

    public static void writeToExcel2015(ArrayList<String> temp, String title){
        Workbook wb = null;
        try {
            FileInputStream fileInputStream = new FileInputStream("file/" + title + ".xlsx");
            wb = WorkbookFactory.create(fileInputStream);
        }catch (Exception e){
            e.printStackTrace();
        }
        Sheet sheet = null;
        sheet = wb.getSheet("综合");
        sheet.getRow(1).createCell(2).setCellValue(temp.get(7));
        sheet.getRow(1).createCell(3).setCellValue("");
        sheet.getRow(2).createCell(2).setCellValue(temp.get(8));
        sheet.getRow(2).createCell(3).setCellValue("");
        sheet.getRow(3).createCell(2).setCellValue("");
        sheet.getRow(3).createCell(3).setCellValue("");
        sheet.getRow(4).createCell(2).setCellValue("");
        sheet.getRow(4).createCell(3).setCellValue("");
        sheet.getRow(5).createCell(2).setCellValue(temp.get(10));
        sheet.getRow(5).createCell(3).setCellValue("");
        sheet.getRow(6).createCell(2).setCellValue(temp.get(11));
        sheet.getRow(6).createCell(3).setCellValue("");
        sheet.getRow(7).createCell(2).setCellValue(temp.get(12));
        sheet.getRow(7).createCell(3).setCellValue("");
        sheet.getRow(8).createCell(2).setCellValue(temp.get(13) + per);
        sheet.getRow(8).createCell(3).setCellValue("");

        sheet.getRow(11).createCell(2).setCellValue(temp.get(16));
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(17) + per);
        sheet.getRow(12).createCell(2).setCellValue("");
        sheet.getRow(12).createCell(3).setCellValue("");
        sheet.getRow(13).createCell(2).setCellValue("");
        sheet.getRow(13).createCell(3).setCellValue("");
        sheet.getRow(14).createCell(2).setCellValue(temp.get(19));
        sheet.getRow(14).createCell(3).setCellValue(add + temp.get(20) + per);
        sheet.getRow(15).createCell(2).setCellValue(temp.get(21));
        sheet.getRow(15).createCell(3).setCellValue(add + temp.get(22) + per);
        sheet.getRow(16).createCell(2).setCellValue(temp.get(23));
        sheet.getRow(16).createCell(3).setCellValue(add + temp.get(24) + per);

        sheet.getRow(20).createCell(2).setCellValue(temp.get(32) + bilyuan);
        sheet.getRow(20).createCell(3).setCellValue(add + temp.get(33) + per);
        sheet.getRow(21).createCell(2).setCellValue("");

        sheet.getRow(30).createCell(2).setCellValue("");
        sheet.getRow(31).createCell(2).setCellValue("");
        sheet.getRow(32).createCell(2).setCellValue("");


        sheet.getRow(40).createCell(2).setCellValue("");


        sheet = wb.getSheet("农业");
        sheet.getRow(1).createCell(2).setCellValue(temp.get(47));
        sheet.getRow(1).createCell(3).setCellValue(add + temp.get(49) + per);
        sheet.getRow(2).createCell(2).setCellValue(temp.get(50));
        sheet.getRow(2).createCell(3).setCellValue(add + temp.get(52) + per);
        sheet.getRow(3).createCell(2).setCellValue(temp.get(53));
        sheet.getRow(3).createCell(3).setCellValue(mis + temp.get(55) + per);
        sheet.getRow(4).createCell(2).setCellValue(temp.get(56));
        sheet.getRow(4).createCell(3).setCellValue(add + temp.get(58) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(59) + milton);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(61) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(62) + milton);
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(64) + per);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(65) + milton);
        sheet.getRow(12).createCell(3).setCellValue(add + temp.get(67) + per);

        sheet.getRow(20).createCell(2).setCellValue(temp.get(68) + milton);
        sheet.getRow(20).createCell(3).setCellValue(add + temp.get(70) + per);


        sheet = wb.getSheet("工业");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(79) + bilyuan);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(81) + hu);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(93) + bilyuan);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(95) + bilyuan);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(97) + bilyuan);


        sheet = wb.getSheet("投资");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(132) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(133) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(138) + bilyuan);
        sheet.getRow(10).createCell(3).setCellValue(add+ temp.get(139) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(140) + bilyuan);
        sheet.getRow(11).createCell(3).setCellValue(add+ temp.get(141) + per);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(144) + bilyuan);
        sheet.getRow(12).createCell(3).setCellValue(add+ temp.get(145) + per);


        sheet = wb.getSheet("贸易");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(159) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(160) + per);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(161) + bilyuan);
        sheet.getRow(1).createCell(3).setCellValue(add + temp.get(162) + per);
        sheet.getRow(2).createCell(2).setCellValue(temp.get(163) + bilyuan);
        sheet.getRow(2).createCell(3).setCellValue(add + temp.get(164) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(171) + bildollar);
        sheet.getRow(10).createCell(3).setCellValue(mis + temp.get(172) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(178) + bildollar);
        sheet.getRow(11).createCell(3).setCellValue(mis + temp.get(179) + per);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(173) + bildollar);
        sheet.getRow(12).createCell(3).setCellValue(mis + temp.get (174) + per);


        sheet = wb.getSheet("交通");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(184) + km);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(189) + km);
        sheet.getRow(2).createCell(2).setCellValue(temp.get(191) + milton);
        sheet.getRow(3).createCell(2).setCellValue(temp.get(193) + milpeo);
        sheet.getRow(4).createCell(2).setCellValue(temp.get(195) + "万人");
        sheet.getRow(5).createCell(2).setCellValue(temp.get(196) + "万吨");

        sheet.getRow(10).createCell(2).setCellValue(temp.get(205) + bilyuan);
        sheet.getRow(10).createCell(3).setCellValue("");
        sheet.getRow(11).createCell(2).setCellValue(temp.get(206) + milhu);
        sheet.getRow(11).createCell(3).setCellValue("");
        sheet.getRow(12).createCell(2).setCellValue(temp.get(207) + milhu);
        sheet.getRow(12).createCell(3).setCellValue("");


        sheet = wb.getSheet("金融");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(217) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(218) + per);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(221) + bilyuan);
        sheet.getRow(1).createCell(3).setCellValue(mis + temp.get(222) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(246) + bilyuan);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(247) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(248) + bilyuan);
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(249) + per);


        sheet = wb.getSheet("教育");
        sheet.getRow(0).createCell(1).setCellValue(temp.get(260) + suo);
        sheet.getRow(0).createCell(3).setCellValue(temp.get(261) + milpeo);
        sheet.getRow(0).createCell(5).setCellValue(temp.get(262) + milpeo);


        sheet.getRow(6).createCell(1).setCellValue(temp.get(263) + suo);
        sheet.getRow(6).createCell(2).setCellValue(temp.get(264) + milpeo);
        sheet.getRow(6).createCell(3).setCellValue(temp.get(265) + milpeo);
        sheet.getRow(7).createCell(1).setCellValue(temp.get(266) + suo);
        sheet.getRow(7).createCell(2).setCellValue(temp.get(267) + milpeo);
        sheet.getRow(7).createCell(3).setCellValue(temp.get(268) + milpeo);
        sheet.getRow(8).createCell(1).setCellValue(temp.get(269) + suo);
        sheet.getRow(8).createCell(2).setCellValue(temp.get(270) + milpeo);
        sheet.getRow(8).createCell(3).setCellValue(temp.get(271) + milpeo);
        sheet.getRow(9).createCell(1).setCellValue(temp.get(272) + suo);
        sheet.getRow(9).createCell(2).setCellValue(temp.get(273) + milpeo);
        sheet.getRow(9).createCell(3).setCellValue(temp.get(274) + milpeo);
        sheet.getRow(10).createCell(1).setCellValue(temp.get(275) + suo);
        sheet.getRow(10).createCell(2).setCellValue(temp.get(276) + milpeo);
        sheet.getRow(10).createCell(3).setCellValue(temp.get(277) + milpeo);


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

    public static void writeToExcel2016(ArrayList<String> temp, String title){
        Workbook wb = null;
        try {
            FileInputStream fileInputStream = new FileInputStream("file/" + title + ".xlsx");
            wb = WorkbookFactory.create(fileInputStream);
        }catch (Exception e){
            e.printStackTrace();
        }
        Sheet sheet = null;
        sheet = wb.getSheet("综合");
        sheet.getRow(1).createCell(2).setCellValue(temp.get(7));
        sheet.getRow(1).createCell(3).setCellValue("");
        sheet.getRow(2).createCell(2).setCellValue(temp.get(8));
        sheet.getRow(2).createCell(3).setCellValue("");
        sheet.getRow(3).createCell(2).setCellValue(temp.get(9));
        sheet.getRow(3).createCell(3).setCellValue("");
        sheet.getRow(4).createCell(2).setCellValue(temp.get(10));
        sheet.getRow(4).createCell(3).setCellValue("");
        sheet.getRow(5).createCell(2).setCellValue(temp.get(11));
        sheet.getRow(5).createCell(3).setCellValue("");
        sheet.getRow(6).createCell(2).setCellValue(temp.get(12));
        sheet.getRow(6).createCell(3).setCellValue("");
        sheet.getRow(7).createCell(2).setCellValue(temp.get(13));
        sheet.getRow(7).createCell(3).setCellValue("");
        sheet.getRow(8).createCell(2).setCellValue(temp.get(14) + per);
        sheet.getRow(8).createCell(3).setCellValue("");

        sheet.getRow(11).createCell(2).setCellValue(temp.get(15));
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(16) + per);
        sheet.getRow(12).createCell(2).setCellValue("");
        sheet.getRow(12).createCell(3).setCellValue("");
        sheet.getRow(13).createCell(2).setCellValue(temp.get(18));
        sheet.getRow(13).createCell(3).setCellValue("");
        sheet.getRow(14).createCell(2).setCellValue(temp.get(19));
        sheet.getRow(14).createCell(3).setCellValue(add + temp.get(20) + per);
        sheet.getRow(15).createCell(2).setCellValue(temp.get(21));
        sheet.getRow(15).createCell(3).setCellValue(add + temp.get(22) + per);
        sheet.getRow(16).createCell(2).setCellValue(temp.get(23));
        sheet.getRow(16).createCell(3).setCellValue(add + temp.get(24) + per);

        sheet.getRow(20).createCell(2).setCellValue(temp.get(28) + bilyuan);
        sheet.getRow(20).createCell(3).setCellValue(add + temp.get(29) + per);
        sheet.getRow(21).createCell(2).setCellValue("");

        sheet.getRow(30).createCell(2).setCellValue("");
        sheet.getRow(31).createCell(2).setCellValue("");
        sheet.getRow(32).createCell(2).setCellValue("");


        sheet.getRow(40).createCell(2).setCellValue("");


        sheet = wb.getSheet("农业");
        sheet.getRow(1).createCell(2).setCellValue(temp.get(39));
        sheet.getRow(1).createCell(3).setCellValue(mis + temp.get(41) + per);
        sheet.getRow(2).createCell(2).setCellValue(temp.get(42));
        sheet.getRow(2).createCell(3).setCellValue("");
        sheet.getRow(3).createCell(2).setCellValue(temp.get(44));
        sheet.getRow(3).createCell(3).setCellValue(mis + temp.get(46) + per);
        sheet.getRow(4).createCell(2).setCellValue(temp.get(47));
        sheet.getRow(4).createCell(3).setCellValue(add + temp.get(49) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(50) + milton);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(52) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(53) + milton);
        sheet.getRow(11).createCell(3).setCellValue("");
        sheet.getRow(12).createCell(2).setCellValue(temp.get(55) + milton);
        sheet.getRow(12).createCell(3).setCellValue(add + temp.get(57) + per);

        sheet.getRow(20).createCell(2).setCellValue(temp.get(58) + milton);
        sheet.getRow(20).createCell(3).setCellValue(mis + temp.get(60) + per);


        sheet = wb.getSheet("工业");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(67) + bilyuan);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(69) + hu);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(79) + bilyuan);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(81) + bilyuan);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(83) + bilyuan);


        sheet = wb.getSheet("投资");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(129) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(130) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(133) + bilyuan);
        sheet.getRow(10).createCell(3).setCellValue(add+ temp.get(134) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(135) + bilyuan);
        sheet.getRow(11).createCell(3).setCellValue(add+ temp.get(136) + per);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(139) + bilyuan);
        sheet.getRow(12).createCell(3).setCellValue(add+ temp.get(140) + per);


        sheet = wb.getSheet("贸易");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(154) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(155) + per);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(156) + bilyuan);
        sheet.getRow(1).createCell(3).setCellValue(add + temp.get(157) + per);
        sheet.getRow(2).createCell(2).setCellValue(temp.get(158) + bilyuan);
        sheet.getRow(2).createCell(3).setCellValue(add + temp.get(159) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(164) + bildollar);
        sheet.getRow(10).createCell(3).setCellValue(mis + temp.get(165) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(171) + bildollar);
        sheet.getRow(11).createCell(3).setCellValue(mis + temp.get(172) + per);
        sheet.getRow(12).createCell(2).setCellValue(temp.get(166) + bildollar);
        sheet.getRow(12).createCell(3).setCellValue(mis + temp.get (167) + per);


        sheet = wb.getSheet("交通");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(175) + km);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(180) + km);
        sheet.getRow(2).createCell(2).setCellValue(temp.get(182) + milton);
        sheet.getRow(3).createCell(2).setCellValue(temp.get(184) + milpeo);
        sheet.getRow(4).createCell(2).setCellValue(temp.get(186) + "万人");
        sheet.getRow(5).createCell(2).setCellValue(temp.get(187) + "万吨");

        sheet.getRow(10).createCell(2).setCellValue(temp.get(189) + bilyuan);
        sheet.getRow(10).createCell(3).setCellValue("");
        sheet.getRow(11).createCell(2).setCellValue(temp.get(190) + milhu);
        sheet.getRow(11).createCell(3).setCellValue("");
        sheet.getRow(12).createCell(2).setCellValue(temp.get(191) + milhu);
        sheet.getRow(12).createCell(3).setCellValue("");


        sheet = wb.getSheet("金融");
        sheet.getRow(0).createCell(2).setCellValue(temp.get(203) + bilyuan);
        sheet.getRow(0).createCell(3).setCellValue(add + temp.get(204) + per);
        sheet.getRow(1).createCell(2).setCellValue(temp.get(207) + bilyuan);
        sheet.getRow(1).createCell(3).setCellValue(add + temp.get(208) + per);

        sheet.getRow(10).createCell(2).setCellValue(temp.get(225) + bilyuan);
        sheet.getRow(10).createCell(3).setCellValue(add + temp.get(226) + per);
        sheet.getRow(11).createCell(2).setCellValue(temp.get(231) + bilyuan);
        sheet.getRow(11).createCell(3).setCellValue(add + temp.get(232) + per);


        sheet = wb.getSheet("教育");
        sheet.getRow(0).createCell(1).setCellValue(temp.get(243) + suo);
        sheet.getRow(0).createCell(3).setCellValue(temp.get(244) + milpeo);
        sheet.getRow(0).createCell(5).setCellValue(temp.get(245) + milpeo);


        sheet.getRow(6).createCell(1).setCellValue(temp.get(246) + suo);
        sheet.getRow(6).createCell(2).setCellValue(temp.get(247) + milpeo);
        sheet.getRow(6).createCell(3).setCellValue(temp.get(248) + milpeo);
        sheet.getRow(7).createCell(1).setCellValue(temp.get(249) + suo);
        sheet.getRow(7).createCell(2).setCellValue(temp.get(250) + milpeo);
        sheet.getRow(7).createCell(3).setCellValue(temp.get(251) + milpeo);
        sheet.getRow(8).createCell(1).setCellValue(temp.get(252) + suo);
        sheet.getRow(8).createCell(2).setCellValue(temp.get(253) + milpeo);
        sheet.getRow(8).createCell(3).setCellValue(temp.get(254) + milpeo);
        sheet.getRow(9).createCell(1).setCellValue(temp.get(255) + suo);
        sheet.getRow(9).createCell(2).setCellValue(temp.get(256) + milpeo);
        sheet.getRow(9).createCell(3).setCellValue(temp.get(257) + milpeo);
        sheet.getRow(10).createCell(1).setCellValue(temp.get(258) + suo);
        sheet.getRow(10).createCell(2).setCellValue(temp.get(259) + milpeo);
        sheet.getRow(10).createCell(3).setCellValue(temp.get(260) + milpeo);


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
}
