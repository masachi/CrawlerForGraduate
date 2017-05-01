package frame;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.jfree.chart.JFreeChart;

import javax.swing.*;
import java.awt.*;
import java.io.FileInputStream;

/**
 * Created by sdlds on 2017/5/1.
 */
public class PeopleEconomyFrame extends JFrame {
    private Workbook wb06 = null;
    private Workbook wb08 = null;
    private Workbook wb09 = null;
    private Workbook wb10 = null;
    private Workbook wb11 = null;
    private Workbook wb12 = null;
    private Workbook wb13 = null;
    private Workbook wb14 = null;
    private Workbook wb15 = null;
    private Workbook wb16 = null;

    public PeopleEconomyFrame(){
        super();
        setTitle("综合图表");
        setBounds(200,200,800,600);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        JTabbedPane tab = new JTabbedPane(JTabbedPane.TOP);

        tab.addTab("年末总户数", null,null,"点击查看年末总户数");
        tab.addTab("城镇/乡村人口", null,null,"点击查看城镇/乡村人口");
        tab.addTab("出生/死亡人口", null,null,"点击查看出生/死亡人口");
        tab.addTab("城镇化率", null,null,"点击查看城镇化率");
        tab.addTab("GDP", null,null,"点击查看GDP");
        tab.addTab("人均GDP", null,null,"点击查看人均GDP");
        tab.addTab("产业增加值", null,null,"点击查看产业增加值");
        tab.addTab("民营经济增加值", null,null,"点击查看民营经济增加值");
        




        tab.setTabLayoutPolicy(JTabbedPane.WRAP_TAB_LAYOUT);
        getContentPane().add(tab, BorderLayout.CENTER);
    }

    private void ReadExcel(){
        try {
            FileInputStream fileInputStream06 = new FileInputStream("file/" + "2006年德阳市国民经济和社会发展统计公报" + ".xlsx");
            FileInputStream fileInputStream08 = new FileInputStream("file/" + "2008年德阳市国民经济和社会发展统计公报" + ".xlsx");
            FileInputStream fileInputStream09 = new FileInputStream("file/" + "2009年德阳市国民经济和社会发展统计公报" + ".xlsx");
            FileInputStream fileInputStream10 = new FileInputStream("file/" + "2010年德阳市国民经济和社会发展统计公报" + ".xlsx");
            FileInputStream fileInputStream11 = new FileInputStream("file/" + "2011年德阳市国民经济和社会发展统计公报" + ".xlsx");
            FileInputStream fileInputStream12 = new FileInputStream("file/" + "2012年德阳市国民经济和社会发展统计公报" + ".xlsx");
            FileInputStream fileInputStream13 = new FileInputStream("file/" + "2013年德阳市国民经济和社会发展统计公报" + ".xlsx");
            FileInputStream fileInputStream14 = new FileInputStream("file/" + "2014年德阳市国民经济和社会发展统计公报" + ".xlsx");
            FileInputStream fileInputStream15 = new FileInputStream("file/" + "2015年德阳市国民经济和社会发展统计公报" + ".xlsx");
            FileInputStream fileInputStream16 = new FileInputStream("file/" + "2016年德阳市国民经济和社会发展统计公报" + ".xlsx");

            wb06 = WorkbookFactory.create(fileInputStream06);
            wb08 = WorkbookFactory.create(fileInputStream08);
            wb09 = WorkbookFactory.create(fileInputStream09);
            wb10 = WorkbookFactory.create(fileInputStream10);
            wb11 = WorkbookFactory.create(fileInputStream11);
            wb12 = WorkbookFactory.create(fileInputStream12);
            wb13 = WorkbookFactory.create(fileInputStream13);
            wb14 = WorkbookFactory.create(fileInputStream14);
            wb15 = WorkbookFactory.create(fileInputStream15);
            wb16 = WorkbookFactory.create(fileInputStream16);

        }catch (Exception e){
            e.printStackTrace();
        }
    }

    private JFreeChart createChartA(){

    }

    private JFreeChart createChartB(){

    }

    private JFreeChart createChartC(){

    }

    private JFreeChart createChartD(){

    }

    private JFreeChart createChartE(){

    }

    private JFreeChart createChartF(){

    }

    private JFreeChart createChartG(){

    }

    private JFreeChart createChartH(){

    }

    public static void main(String[] args){
        PeopleEconomyFrame frame = new PeopleEconomyFrame();
        frame.setVisible(true);
    }
}
