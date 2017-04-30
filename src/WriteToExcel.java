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

        sheet.getRow(20).createCell(2).setCellValue(temp.get(40));
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
        sheet.getRow(2).createCell(2).setCellValue(temp.get(264) + "亿吨公里");
        sheet.getRow(3).createCell(2).setCellValue(temp.get(265) + "亿人公里");
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

        sheet.getRow(20).createCell(2).setCellValue(temp.get(28));
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
        sheet.getRow(2).createCell(2).setCellValue(temp.get(187) + "亿吨公里");
        sheet.getRow(3).createCell(2).setCellValue(temp.get(189) + "亿人公里");
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

        sheet.getRow(20).createCell(2).setCellValue(temp.get(21));
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
        sheet.getRow(2).createCell(2).setCellValue(temp.get(169) + "亿吨公里");
        sheet.getRow(3).createCell(2).setCellValue(temp.get(171) + "亿人公里");
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

    }

    public static void writeToExcel2012(ArrayList<String> temp, String title){

    }

    public static void writeToExcel2013(ArrayList<String> temp, String title){

    }

    public static void writeToExcel2014(ArrayList<String> temp, String title){

    }

    public static void writeToExcel2015(ArrayList<String> temp, String title){

    }

    public static void writeToExcel2016(ArrayList<String> temp, String title){

    }
}
