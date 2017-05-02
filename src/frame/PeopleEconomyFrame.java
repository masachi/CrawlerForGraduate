package frame;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.labels.ItemLabelAnchor;
import org.jfree.chart.labels.ItemLabelPosition;
import org.jfree.chart.labels.StandardCategoryItemLabelGenerator;
import org.jfree.chart.plot.CategoryPlot;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.renderer.category.BarRenderer3D;
import org.jfree.data.category.CategoryDataset;
import org.jfree.data.general.DatasetUtilities;
import org.jfree.ui.TextAnchor;

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

    public PeopleEconomyFrame() {
        super();
        ReadExcel();
        setTitle("综合图表");
        setBounds(200, 200, 800, 600);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        JTabbedPane tab = new JTabbedPane(JTabbedPane.TOP);

        tab.addTab("年末总户数", null, createChartA(), "点击查看年末总户数");
        tab.addTab("户籍总人口", null, createChartB(), "点击查看户籍总人口");
        tab.addTab("城镇/乡村人口", null, createChartC(), "点击查看城镇/乡村人口");
        tab.addTab("出生/死亡人口", null, createChartD(), "点击查看出生/死亡人口");
        tab.addTab("城镇化率", null, createChartE(), "点击查看城镇化率");
        tab.addTab("GDP", null, createChartF(), "点击查看GDP");
        tab.addTab("人均GDP", null, createChartG(), "点击查看人均GDP");
        tab.addTab("产业增加值", null, createChartH(), "点击查看产业增加值");
        tab.addTab("民营经济增加值", null, createChartI(), "点击查看民营经济增加值");


        tab.setTabLayoutPolicy(JTabbedPane.WRAP_TAB_LAYOUT);
        getContentPane().add(tab, BorderLayout.CENTER);
    }

    private void ReadExcel() {
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

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private ChartPanel createChartA() {
        double[][] data = new double[][]{
                {Double.parseDouble(wb06.getSheet("综合").getRow(1).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb08.getSheet("综合").getRow(1).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb09.getSheet("综合").getRow(1).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb11.getSheet("综合").getRow(1).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb12.getSheet("综合").getRow(1).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb13.getSheet("综合").getRow(1).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb14.getSheet("综合").getRow(1).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb15.getSheet("综合").getRow(1).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb16.getSheet("综合").getRow(1).getCell(2).getStringCellValue())}
        };
        String[] rowKeys = {"2006", "2008", "2009", "2011", "2012", "2013", "2014", "2015", "2016"};
        String[] columnKeys = {""};
        CategoryDataset dataset = DatasetUtilities.createCategoryDataset(rowKeys, columnKeys, data);
        JFreeChart chart = ChartFactory.createBarChart3D("年末总户数统计图", "年份", "总户数（万户）", dataset, PlotOrientation.VERTICAL, true, true, false);

        CategoryPlot plot = chart.getCategoryPlot();
        BarRenderer3D renderer = new BarRenderer3D();
        renderer.setBaseItemLabelGenerator(new StandardCategoryItemLabelGenerator());
        renderer.setBaseItemLabelsVisible(true);
        //默认的数字显示在柱子中，通过如下两句可调整数字的显示
        //注意：此句很关键，若无此句，那数字的显示会被覆盖，给人数字没有显示出来的问题
        renderer.setBasePositiveItemLabelPosition(new ItemLabelPosition(ItemLabelAnchor.OUTSIDE12, TextAnchor.BASELINE_LEFT));
        renderer.setItemLabelAnchorOffset(10D);
        plot.setRenderer(renderer);

        return new ChartPanel(chart);
    }

    private ChartPanel createChartB(){
        double[][] data = new double[][] {
                {Double.parseDouble(wb08.getSheet("综合").getRow(2).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb09.getSheet("综合").getRow(2).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb11.getSheet("综合").getRow(2).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb12.getSheet("综合").getRow(2).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb13.getSheet("综合").getRow(2).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb14.getSheet("综合").getRow(2).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb15.getSheet("综合").getRow(2).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb16.getSheet("综合").getRow(2).getCell(2).getStringCellValue())}
        };
        String[] rowKeys = {"2008", "2009", "2011", "2012", "2013", "2014", "2015", "2016"};
        String[] columnKeys = {""};

        CategoryDataset dataset = DatasetUtilities.createCategoryDataset(rowKeys, columnKeys, data);
        JFreeChart chart = ChartFactory.createBarChart3D("户籍总人口统计图", "年份", "总人口（万人）", dataset, PlotOrientation.VERTICAL, true, true, false);

        CategoryPlot plot = chart.getCategoryPlot();
        BarRenderer3D renderer = new BarRenderer3D();
        renderer.setBaseItemLabelGenerator(new StandardCategoryItemLabelGenerator());
        renderer.setBaseItemLabelsVisible(true);
        //默认的数字显示在柱子中，通过如下两句可调整数字的显示
        //注意：此句很关键，若无此句，那数字的显示会被覆盖，给人数字没有显示出来的问题
        renderer.setBasePositiveItemLabelPosition(new ItemLabelPosition(ItemLabelAnchor.OUTSIDE12, TextAnchor.BASELINE_LEFT));
        renderer.setItemLabelAnchorOffset(10D);
        plot.setRenderer(renderer);

        return new ChartPanel(chart);
    }

    private ChartPanel createChartC(){
        double[][] data = new double[][]{
                {
                        Double.parseDouble(wb06.getSheet("综合").getRow(3).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb08.getSheet("综合").getRow(3).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb09.getSheet("综合").getRow(3).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb16.getSheet("综合").getRow(3).getCell(2).getStringCellValue())
                },
                {
                        Double.parseDouble(wb06.getSheet("综合").getRow(4).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb08.getSheet("综合").getRow(4).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb09.getSheet("综合").getRow(4).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb16.getSheet("综合").getRow(4).getCell(2).getStringCellValue())
                }
        };
        String[] columnKeys = {"2006", "2008", "2009", "2016"};
        String[] rowKeys = {"城镇","乡村"};

        CategoryDataset dataset = DatasetUtilities.createCategoryDataset(rowKeys, columnKeys, data);
        JFreeChart chart = ChartFactory.createBarChart3D("城镇／乡村人口统计图", "年份", "人口（万人）", dataset, PlotOrientation.VERTICAL, true, true, false);

        CategoryPlot plot = chart.getCategoryPlot();
        BarRenderer3D renderer = new BarRenderer3D();
        renderer.setBaseItemLabelGenerator(new StandardCategoryItemLabelGenerator());
        renderer.setBaseItemLabelsVisible(true);
        //默认的数字显示在柱子中，通过如下两句可调整数字的显示
        //注意：此句很关键，若无此句，那数字的显示会被覆盖，给人数字没有显示出来的问题
        renderer.setBasePositiveItemLabelPosition(new ItemLabelPosition(ItemLabelAnchor.OUTSIDE12, TextAnchor.BASELINE_LEFT));
        renderer.setItemLabelAnchorOffset(10D);
        plot.setRenderer(renderer);

        return new ChartPanel(chart);
    }

    private ChartPanel createChartD(){
        double[][] data = new double[][] {
                {
                        Double.parseDouble(wb06.getSheet("综合").getRow(5).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb08.getSheet("综合").getRow(5).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb09.getSheet("综合").getRow(5).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb11.getSheet("综合").getRow(5).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb12.getSheet("综合").getRow(5).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb13.getSheet("综合").getRow(5).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb14.getSheet("综合").getRow(5).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb15.getSheet("综合").getRow(5).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb16.getSheet("综合").getRow(5).getCell(2).getStringCellValue()),
                },
                {
                        Double.parseDouble(wb06.getSheet("综合").getRow(6).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb08.getSheet("综合").getRow(6).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb09.getSheet("综合").getRow(6).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb11.getSheet("综合").getRow(6).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb12.getSheet("综合").getRow(6).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb13.getSheet("综合").getRow(6).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb14.getSheet("综合").getRow(6).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb15.getSheet("综合").getRow(6).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb16.getSheet("综合").getRow(6).getCell(2).getStringCellValue()),
                }
        };
        String[] columnKeys = {"2006", "2008", "2009", "2011", "2012", "2013", "2014", "2015", "2016"};
        String[] rowKeys = {"出生","死亡"};

        CategoryDataset dataset = DatasetUtilities.createCategoryDataset(rowKeys, columnKeys, data);
        JFreeChart chart = ChartFactory.createBarChart3D("出生／死亡人口统计图", "年份", "人口（万人）", dataset, PlotOrientation.VERTICAL, true, true, false);

        CategoryPlot plot = chart.getCategoryPlot();
        BarRenderer3D renderer = new BarRenderer3D();
        renderer.setBaseItemLabelGenerator(new StandardCategoryItemLabelGenerator());
        renderer.setBaseItemLabelsVisible(true);
        //默认的数字显示在柱子中，通过如下两句可调整数字的显示
        //注意：此句很关键，若无此句，那数字的显示会被覆盖，给人数字没有显示出来的问题
        renderer.setBasePositiveItemLabelPosition(new ItemLabelPosition(ItemLabelAnchor.OUTSIDE12, TextAnchor.BASELINE_LEFT));
        renderer.setItemLabelAnchorOffset(10D);
        plot.setRenderer(renderer);

        return new ChartPanel(chart);
    }

    private ChartPanel createChartE(){
        double[][] data = new double[][]{
                {Double.parseDouble(wb06.getSheet("综合").getRow(8).getCell(2).getStringCellValue().replace("%",""))},
                {Double.parseDouble(wb08.getSheet("综合").getRow(8).getCell(2).getStringCellValue().replace("%",""))},
                {Double.parseDouble(wb09.getSheet("综合").getRow(8).getCell(2).getStringCellValue().replace("%",""))},
                {Double.parseDouble(wb11.getSheet("综合").getRow(8).getCell(2).getStringCellValue().replace("%",""))},
                {Double.parseDouble(wb12.getSheet("综合").getRow(8).getCell(2).getStringCellValue().replace("%",""))},
                {Double.parseDouble(wb13.getSheet("综合").getRow(8).getCell(2).getStringCellValue().replace("%",""))},
                {Double.parseDouble(wb14.getSheet("综合").getRow(8).getCell(2).getStringCellValue().replace("%",""))},
                {Double.parseDouble(wb15.getSheet("综合").getRow(8).getCell(2).getStringCellValue().replace("%",""))},
                {Double.parseDouble(wb16.getSheet("综合").getRow(8).getCell(2).getStringCellValue().replace("%",""))}
        };
        String[] rowKeys = {"2006", "2008", "2009", "2011", "2012", "2013", "2014", "2015", "2016"};
        String[] columnKeys = {""};
        CategoryDataset dataset = DatasetUtilities.createCategoryDataset(rowKeys, columnKeys, data);
        JFreeChart chart = ChartFactory.createBarChart3D("城镇化率统计图", "年份", "城镇化率（%）", dataset, PlotOrientation.VERTICAL, true, true, false);

        CategoryPlot plot = chart.getCategoryPlot();
        BarRenderer3D renderer = new BarRenderer3D();
        renderer.setBaseItemLabelGenerator(new StandardCategoryItemLabelGenerator());
        renderer.setBaseItemLabelsVisible(true);
        //默认的数字显示在柱子中，通过如下两句可调整数字的显示
        //注意：此句很关键，若无此句，那数字的显示会被覆盖，给人数字没有显示出来的问题
        renderer.setBasePositiveItemLabelPosition(new ItemLabelPosition(ItemLabelAnchor.OUTSIDE12, TextAnchor.BASELINE_LEFT));
        renderer.setItemLabelAnchorOffset(10D);
        plot.setRenderer(renderer);

        return new ChartPanel(chart);
    }

    private ChartPanel createChartF(){
        double[][] data = new double[][]{
                {Double.parseDouble(wb06.getSheet("综合").getRow(11).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb08.getSheet("综合").getRow(11).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb09.getSheet("综合").getRow(11).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb10.getSheet("综合").getRow(11).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb11.getSheet("综合").getRow(11).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb12.getSheet("综合").getRow(11).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb13.getSheet("综合").getRow(11).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb14.getSheet("综合").getRow(11).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb15.getSheet("综合").getRow(11).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb16.getSheet("综合").getRow(11).getCell(2).getStringCellValue())}
        };
        String[] rowKeys = {"2006", "2008", "2010", "2009", "2011", "2012", "2013", "2014", "2015", "2016"};
        String[] columnKeys = {""};
        CategoryDataset dataset = DatasetUtilities.createCategoryDataset(rowKeys, columnKeys, data);
        JFreeChart chart = ChartFactory.createBarChart3D("GDP统计图", "年份", "GDP（亿元）", dataset, PlotOrientation.VERTICAL, true, true, false);

        CategoryPlot plot = chart.getCategoryPlot();
        BarRenderer3D renderer = new BarRenderer3D();
        renderer.setBaseItemLabelGenerator(new StandardCategoryItemLabelGenerator());
        renderer.setBaseItemLabelsVisible(true);
        //默认的数字显示在柱子中，通过如下两句可调整数字的显示
        //注意：此句很关键，若无此句，那数字的显示会被覆盖，给人数字没有显示出来的问题
        renderer.setBasePositiveItemLabelPosition(new ItemLabelPosition(ItemLabelAnchor.OUTSIDE12, TextAnchor.BASELINE_LEFT));
        renderer.setItemLabelAnchorOffset(10D);
        plot.setRenderer(renderer);

        return new ChartPanel(chart);
    }

    private ChartPanel createChartG(){
        double[][] data = new double[][]{
                {Double.parseDouble(wb06.getSheet("综合").getRow(13).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb08.getSheet("综合").getRow(13).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb09.getSheet("综合").getRow(13).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb11.getSheet("综合").getRow(13).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb12.getSheet("综合").getRow(13).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb13.getSheet("综合").getRow(13).getCell(2).getStringCellValue())},
                {Double.parseDouble(wb16.getSheet("综合").getRow(13).getCell(2).getStringCellValue())}
        };
        String[] rowKeys = {"2006", "2008", "2009", "2011", "2012", "2013", "2016"};
        String[] columnKeys = {""};
        CategoryDataset dataset = DatasetUtilities.createCategoryDataset(rowKeys, columnKeys, data);
        JFreeChart chart = ChartFactory.createBarChart3D("人均GDP统计图", "年份", "人均GDP（元）", dataset, PlotOrientation.VERTICAL, true, true, false);

        CategoryPlot plot = chart.getCategoryPlot();
        BarRenderer3D renderer = new BarRenderer3D();
        renderer.setBaseItemLabelGenerator(new StandardCategoryItemLabelGenerator());
        renderer.setBaseItemLabelsVisible(true);
        //默认的数字显示在柱子中，通过如下两句可调整数字的显示
        //注意：此句很关键，若无此句，那数字的显示会被覆盖，给人数字没有显示出来的问题
        renderer.setBasePositiveItemLabelPosition(new ItemLabelPosition(ItemLabelAnchor.OUTSIDE12, TextAnchor.BASELINE_LEFT));
        renderer.setItemLabelAnchorOffset(10D);
        plot.setRenderer(renderer);

        return new ChartPanel(chart);
    }

    private ChartPanel createChartH(){
        double[][] data = new double[][] {
                {
                        Double.parseDouble(wb06.getSheet("综合").getRow(14).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb08.getSheet("综合").getRow(14).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb09.getSheet("综合").getRow(14).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb10.getSheet("综合").getRow(14).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb13.getSheet("综合").getRow(14).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb14.getSheet("综合").getRow(14).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb15.getSheet("综合").getRow(14).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb16.getSheet("综合").getRow(14).getCell(2).getStringCellValue()),
                },
                {
                        Double.parseDouble(wb06.getSheet("综合").getRow(15).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb08.getSheet("综合").getRow(15).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb09.getSheet("综合").getRow(15).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb10.getSheet("综合").getRow(15).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb13.getSheet("综合").getRow(15).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb14.getSheet("综合").getRow(15).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb15.getSheet("综合").getRow(15).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb16.getSheet("综合").getRow(15).getCell(2).getStringCellValue()),
                },
                {
                        Double.parseDouble(wb06.getSheet("综合").getRow(16).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb08.getSheet("综合").getRow(16).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb09.getSheet("综合").getRow(16).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb10.getSheet("综合").getRow(16).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb13.getSheet("综合").getRow(16).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb14.getSheet("综合").getRow(16).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb15.getSheet("综合").getRow(16).getCell(2).getStringCellValue()),
                        Double.parseDouble(wb16.getSheet("综合").getRow(16).getCell(2).getStringCellValue()),
                }

        };
        String[] columnKeys = {"2006", "2008", "2009", "2010", "2013","2014","2015", "2016"};
        String[] rowKeys = {"第一产业增加值","第二产业增加值","第三产业增加值"};
        CategoryDataset dataset = DatasetUtilities.createCategoryDataset(rowKeys, columnKeys, data);
        JFreeChart chart = ChartFactory.createBarChart3D("产业增加值统计图", "年份", "增加值（亿元）", dataset, PlotOrientation.VERTICAL, true, true, false);

        CategoryPlot plot = chart.getCategoryPlot();
        BarRenderer3D renderer = new BarRenderer3D();
        renderer.setBaseItemLabelGenerator(new StandardCategoryItemLabelGenerator());
        renderer.setBaseItemLabelsVisible(true);
        //默认的数字显示在柱子中，通过如下两句可调整数字的显示
        //注意：此句很关键，若无此句，那数字的显示会被覆盖，给人数字没有显示出来的问题
        renderer.setBasePositiveItemLabelPosition(new ItemLabelPosition(ItemLabelAnchor.OUTSIDE12, TextAnchor.BASELINE_LEFT));
        renderer.setItemLabelAnchorOffset(10D);
        plot.setRenderer(renderer);

        return new ChartPanel(chart);
    }
    private ChartPanel createChartI(){
        double[][] data = new double[][]{
                {Double.parseDouble(wb08.getSheet("综合").getRow(11).getCell(2).getStringCellValue().replace("亿元",""))},
                {Double.parseDouble(wb09.getSheet("综合").getRow(11).getCell(2).getStringCellValue().replace("亿元",""))},
                {Double.parseDouble(wb10.getSheet("综合").getRow(11).getCell(2).getStringCellValue().replace("亿元",""))},
                {Double.parseDouble(wb11.getSheet("综合").getRow(11).getCell(2).getStringCellValue().replace("亿元",""))},
                {Double.parseDouble(wb12.getSheet("综合").getRow(11).getCell(2).getStringCellValue().replace("亿元",""))},
                {Double.parseDouble(wb13.getSheet("综合").getRow(11).getCell(2).getStringCellValue().replace("亿元",""))},
                {Double.parseDouble(wb14.getSheet("综合").getRow(11).getCell(2).getStringCellValue().replace("亿元",""))},
                {Double.parseDouble(wb15.getSheet("综合").getRow(11).getCell(2).getStringCellValue().replace("亿元",""))},
                {Double.parseDouble(wb16.getSheet("综合").getRow(11).getCell(2).getStringCellValue().replace("亿元",""))}
        };
        String[] rowKeys = {"2008", "2010", "2009", "2011", "2012", "2013", "2014", "2015", "2016"};
        String[] columnKeys = {""};
        CategoryDataset dataset = DatasetUtilities.createCategoryDataset(rowKeys, columnKeys, data);
        JFreeChart chart = ChartFactory.createBarChart3D("民营经济增加值统计图", "年份", "经济增加值（亿元）", dataset, PlotOrientation.VERTICAL, true, true, false);

        CategoryPlot plot = chart.getCategoryPlot();
        BarRenderer3D renderer = new BarRenderer3D();
        renderer.setBaseItemLabelGenerator(new StandardCategoryItemLabelGenerator());
        renderer.setBaseItemLabelsVisible(true);
        //默认的数字显示在柱子中，通过如下两句可调整数字的显示
        //注意：此句很关键，若无此句，那数字的显示会被覆盖，给人数字没有显示出来的问题
        renderer.setBasePositiveItemLabelPosition(new ItemLabelPosition(ItemLabelAnchor.OUTSIDE12, TextAnchor.BASELINE_LEFT));
        renderer.setItemLabelAnchorOffset(10D);
        plot.setRenderer(renderer);

        return new ChartPanel(chart);
    }

    public static void main(String[] args) {
        PeopleEconomyFrame frame = new PeopleEconomyFrame();
        frame.setVisible(true);
    }
}
