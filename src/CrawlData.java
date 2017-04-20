import com.gargoylesoftware.htmlunit.BrowserVersion;
import com.gargoylesoftware.htmlunit.NicelyResynchronizingAjaxController;
import com.gargoylesoftware.htmlunit.WebClient;
import com.gargoylesoftware.htmlunit.html.HtmlPage;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by Masachi on 2017/4/18.
 */
public class CrawlData implements Runnable {
    private String title;
    private String url;
    private Document doc = null;
    private static WebClient client;
    private String originPage;
    private Workbook wb = null;

    public CrawlData(String title, String url) {
        this.url = url;
        this.title = title;
    }

    private void getPageFromWeb(String URL) throws Exception {
        client = new WebClient(BrowserVersion.CHROME);
        client.getOptions().setJavaScriptEnabled(true); //启用JS解释器，默认为true
        client.getOptions().setCssEnabled(false); //禁用css支持
//        client.getOptions().setProxyConfig(new ProxyConfig("185.10.17.134",3128));
        client.getCookieManager().setCookiesEnabled(false);
        client.getOptions().setThrowExceptionOnScriptError(false); //js运行错误时，是否抛出异常
        client.getOptions().setThrowExceptionOnFailingStatusCode(false);
        client.getOptions().setTimeout(10000); //设置连接超时时间 ，这里是10S。如果为0，则无限期等待

        client.waitForBackgroundJavaScript(600 * 1000);
        client.setAjaxController(new NicelyResynchronizingAjaxController());

        HtmlPage page = client.getPage(URL);
        client.waitForBackgroundJavaScript(1000 * 3);
        client.setJavaScriptTimeout(0);
        //System.out.println(page);
        String pageXml = page.asXml(); //以xml的形式获取响应文本
        //System.out.println(pageXml);
//        doc = Jsoup.connect(Url).get();
        originPage = pageXml;
        doc = Jsoup.parse(pageXml);
    }

    private void getActualUrl() {
        String actualUrl = MultiThread.baseUrl + doc.getElementById("frame_content").attr("src");
        //System.out.println("actualUrl: "+ actualUrl);

        int year = Integer.parseInt(title.substring(0, 4));
        findData(actualUrl, year);
    }

    private void findData(String urlOld, int year) {
        try {
            ArrayList<String> temp = new ArrayList<>();
            getPageFromWeb(urlOld);
            String totalText = originPage.replaceAll("<[^>]+>", "").replaceAll("\r\n", "").replaceAll(" ", "");
            //System.out.println(totalText);

            Matcher m = Pattern.compile("(\\d+\\.\\d+|\\d+)").matcher(totalText);
            temp.add(String.valueOf(0));
            while(m.find()){
                //System.out.println(m.group());
                temp.add(m.group());
            }



//            Matcher peopleOld = Pattern.compile(".*?年末总户数(\\d+.\\d+|\\d+)万户，户籍总人口(\\d+.\\d+|\\d+)万人。全年出生人口(\\d+.\\d+|\\d+)万人，死亡人口(\\d+.\\d+|\\d+)万人。年末常住人口(\\d+.\\d+|\\d+)万人，其中城镇常住人口(\\d+.\\d+|\\d+)万人，城镇化率(\\d+.\\d+|\\d+)%。.*?").matcher(totalText);
//            Matcher economyOld = Pattern.compile(".*?全年地区生产总值(\\d+.\\d+|\\d+)亿元，按可比价格计算，比上年增长(\\d+.\\d+|\\d+)%，其中，第一产业增加值(\\d+.\\d+|\\d+)亿元，比上年增长(\\d+.\\d+|\\d+)%；第二产业增加值(\\d+.\\d+|\\d+)亿元，增长(\\d+.\\d+|\\d+) %；第三产增加值(\\d+.\\d+|\\d+) 亿元，增长(\\d+.\\d+|\\d+)%。.*?全年非公有制经济实现增加值(\\d+.\\d+|\\d+)亿元，增长(\\d+.\\d+|\\d+)%。非公有制经济占全市经济的比重为(\\d+.\\d+|\\d+)%，比上年增加(\\d+.\\d+|\\d+)个百分点。").matcher(totalText);
//            Matcher agriculturalOld = Pattern.compile("").matcher(totalText);
//            Matcher industryOld = Pattern.compile("").matcher(totalText);
//            Matcher investmentOld = Pattern.compile("").matcher(totalText);
//            Matcher tradeOld = Pattern.compile("").matcher(totalText);
//            Matcher transportOld = Pattern.compile("").matcher(totalText);
//            Matcher ensuranceOld = Pattern.compile("").matcher(totalText);
//            Matcher educationOld = Pattern.compile("").matcher(totalText);
//            Matcher cultureOld = Pattern.compile("").matcher(totalText);
//            Matcher societyOLd = Pattern.compile("").matcher(totalText);
//            Matcher environmentOld = Pattern.compile("").matcher(totalText);totalText

//            while (peopleOld.find()) {
//                System.out.println(peopleOld.group(1));
//                System.out.println(peopleOld.group(2));
//                System.out.println(peopleOld.group(3));
//                System.out.println(peopleOld.group(4));
//                System.out.println(peopleOld.group(5));
//                System.out.println(peopleOld.group(6));
//                System.out.println(peopleOld.group(7));
//            }

//            while(economyOld.find()){
//                System.out.println(economyOld.group(1));
//                System.out.println(economyOld.group(2));
//                System.out.println(economyOld.group(3));
//                System.out.println(economyOld.group(4));
//                System.out.println(economyOld.group(5));
//                System.out.println(economyOld.group(6));
//                System.out.println(economyOld.group(7));
//                System.out.println(economyOld.group(8));
//            }
            switch(year){
                case 2006:
                    WriteToExcel.writeToExcel2006(temp,title);
                    break;
                default:
                    break;
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void initExcel(){
        wb = new XSSFWorkbook();
        initPeopleSheet();
        initAgriculture();
        initIndustry();
        initInvestment();
        initTrade();
        initTransport();
        initEnsurance();
        initEducation();
        OutputExcel();
    }

    private void initPeopleSheet(){
        Sheet peopleSheet = wb.createSheet("综合");
        Row row = null;
        Cell cell = null;
        row = peopleSheet.createRow(0);
        row.createCell(0).setCellValue("人口");
        row.createCell(1).setCellValue("计量单位");
        row.createCell(2).setCellValue("数值");
        row.createCell(3).setCellValue("同比");
        row = peopleSheet.createRow(1);
        row.createCell(0).setCellValue("年末总户数");
        row.createCell(1).setCellValue("万户");
        row = peopleSheet.createRow(2);
        row.createCell(0).setCellValue("户籍总人口");
        row.createCell(1).setCellValue("万人");
        row = peopleSheet.createRow(3);
        row.createCell(0).setCellValue("城镇人口");
        row.createCell(1).setCellValue("万人");
        row = peopleSheet.createRow(4);
        row.createCell(0).setCellValue("乡村人口");
        row.createCell(1).setCellValue("万人");
        row = peopleSheet.createRow(5);
        row.createCell(0).setCellValue("全年出生人口");
        row.createCell(1).setCellValue("万人");
        row = peopleSheet.createRow(6);
        row.createCell(0).setCellValue("死亡人口");
        row.createCell(1).setCellValue("万人");
        row = peopleSheet.createRow(7);
        row.createCell(0).setCellValue("年末常住人口");
        row.createCell(1).setCellValue("万人");
        row = peopleSheet.createRow(8);
        row.createCell(0).setCellValue("城镇化率");

        row = peopleSheet.createRow(10);
        row.createCell(0).setCellValue("国民经济");
        row.createCell(1).setCellValue("计量单位");
        row.createCell(2).setCellValue("数值");
        row.createCell(3).setCellValue("同比");
        row = peopleSheet.createRow(11);
        row.createCell(0).setCellValue("全年地区生产总值");
        row.createCell(1).setCellValue("亿元");
        row = peopleSheet.createRow(12);
        row.createCell(0).setCellValue("经济总量");
        row.createCell(1).setCellValue("亿元");
        row = peopleSheet.createRow(13);
        row.createCell(0).setCellValue("人均GDP ");
        row.createCell(1).setCellValue("元");
        row = peopleSheet.createRow(14);
        row.createCell(0).setCellValue("第一产业增加值");
        row.createCell(1).setCellValue("亿元");
        row = peopleSheet.createRow(15);
        row.createCell(0).setCellValue("第二产业增加值");
        row.createCell(1).setCellValue("亿元");
        row = peopleSheet.createRow(16);
        row.createCell(0).setCellValue("第三产业增加值");
        row.createCell(1).setCellValue("亿元");

        row = peopleSheet.createRow(20);
        row.createCell(0).setCellValue("全年民营经济");
        row = peopleSheet.createRow(21);
        row.createCell(0).setCellValue("民营经济");

        row = peopleSheet.createRow(30);
        row.createCell(0).setCellValue("全年居民消费价格");
        row = peopleSheet.createRow(31);
        row.createCell(0).setCellValue("工业品出厂价格");
        row = peopleSheet.createRow(32);
        row.createCell(0).setCellValue("工业品购进价格");

        row = peopleSheet.createRow(40);
        row.createCell(0).setCellValue("一般公共预算");
    }

    private void initAgriculture(){
        Sheet agriSheet = wb.createSheet("农业");
        Row row = null;
        Cell cell = null;
        row = agriSheet.createRow(0);
        row.createCell(0).setCellValue("播种面积");
        row.createCell(1).setCellValue("计量单位");
        row.createCell(2).setCellValue("数值");
        row.createCell(3).setCellValue("同比");
        row = agriSheet.createRow(1);
        row.createCell(0).setCellValue("全年农作物播种面积");
        row.createCell(1).setCellValue("万公顷");
        row = agriSheet.createRow(2);
        row.createCell(0).setCellValue("粮食作物播种面积");
        row.createCell(1).setCellValue("万公顷");
        row = agriSheet.createRow(3);
        row.createCell(0).setCellValue("油料作物播种面积");
        row.createCell(1).setCellValue("万公顷");
        row = agriSheet.createRow(4);
        row.createCell(0).setCellValue("蔬菜及食用菌播种面积");
        row.createCell(1).setCellValue("万公顷");

        row = agriSheet.createRow(10);
        row.createCell(0).setCellValue("全年粮食总产量");
        row = agriSheet.createRow(11);
        row.createCell(0).setCellValue("油料总产量");
        row = agriSheet.createRow(12);
        row.createCell(0).setCellValue("蔬菜总产量");

        row = agriSheet.createRow(20);
        row.createCell(0).setCellValue("全年肉类总产量");
    }

    private void initIndustry(){
        Sheet industSheet = wb.createSheet("工业");
        Row row = null;
        Cell cell = null;

        row = industSheet.createRow(0);
        row.createCell(0).setCellValue("全部工业增加值");
        row = industSheet.createRow(1);
        row.createCell(0).setCellValue("工业企业");

        row = industSheet.createRow(10);
        row.createCell(0).setCellValue("主营业务收入");
        row = industSheet.createRow(11);
        row.createCell(0).setCellValue("实现利税");
        row = industSheet.createRow(12);
        row.createCell(0).setCellValue("利润总额");
    }

    private void initInvestment(){
        Sheet investSheet = wb.createSheet("投资");
        Row row = null;
        Cell cell = null;

        row = investSheet.createRow(0);
        row.createCell(0).setCellValue("固定资产投资");

        row = investSheet.createRow(10);
        row.createCell(0).setCellValue("第一产业投资");
        row = investSheet.createRow(11);
        row.createCell(0).setCellValue("第二产业投资");
        row = investSheet.createRow(12);
        row.createCell(0).setCellValue("第三产业投资");
    }

    private void initTrade(){
        Sheet tradeSheet = wb.createSheet("贸易");
        Row row = null;
        Cell cell = null;

        row = tradeSheet.createRow(0);
        row.createCell(0).setCellValue("社会消费品零售总额");
        row = tradeSheet.createRow(1);
        row.createCell(0).setCellValue("城镇消费品零售额");
        row = tradeSheet.createRow(2);
        row.createCell(0).setCellValue("乡村消费品零售额");

        row = tradeSheet.createRow(10);
        row.createCell(0).setCellValue("进出口总额");
        row = tradeSheet.createRow(11);
        row.createCell(0).setCellValue("出口额");
        row = tradeSheet.createRow(12);
        row.createCell(0).setCellValue("进口额");
    }

    private void initTransport(){
        Sheet transSheet = wb.createSheet("交通");
        Row row = null;
        Cell cell = null;

        row = transSheet.createRow(0);
        row.createCell(0).setCellValue("公路总里程");
        row = transSheet.createRow(1);
        row.createCell(0).setCellValue("铁路运营里程");
        row = transSheet.createRow(2);
        row.createCell(0).setCellValue("公路货运量");
        row = transSheet.createRow(3);
        row.createCell(0).setCellValue("公路客运量");
        row = transSheet.createRow(4);
        row.createCell(0).setCellValue("铁路客运量");
        row = transSheet.createRow(5);
        row.createCell(0).setCellValue("铁路货运量");

        row = transSheet.createRow(10);
        row.createCell(0).setCellValue("电信业务总量");
        row = transSheet.createRow(11);
        row.createCell(0).setCellValue("固定电话用户");
        row = transSheet.createRow(12);
        row.createCell(0).setCellValue("移动电话用户");
    }

    private void initEnsurance(){
        Sheet ensuresheet = wb.createSheet("金融");
        Row row = null;
        Cell cell = null;

        row = ensuresheet.createRow(0);
        row.createCell(0).setCellValue("金融机构本外币存款余额");
        row = ensuresheet.createRow(1);
        row.createCell(0).setCellValue("金融机构本外币贷款余额");

        row = ensuresheet.createRow(10);
        row.createCell(0).setCellValue("保费收入");
        row = ensuresheet.createRow(11);
        row.createCell(0).setCellValue("各项赔款");
    }

    private void initEducation(){
        Sheet eduSheet = wb.createSheet("教育");
        Row row = null;
        Cell cell = null;

        row = eduSheet.createRow(0);
        row.createCell(0).setCellValue("学校");
        row.createCell(2).setCellValue("教师");
        row.createCell(4).setCellValue("在校学生");
        row = eduSheet.createRow(5);
        row.createCell(1).setCellValue("数量");
        row.createCell(2).setCellValue("招生人数");
        row.createCell(3).setCellValue("在校");
        row = eduSheet.createRow(6);
        row.createCell(0).setCellValue("小学");
        row = eduSheet.createRow(7);
        row.createCell(0).setCellValue("初中");
        row = eduSheet.createRow(8);
        row.createCell(0).setCellValue("高中");
        row = eduSheet.createRow(9);
        row.createCell(0).setCellValue("中等职业学校");
        row = eduSheet.createRow(10);
        row.createCell(0).setCellValue("普通高校");
    }

    private void OutputExcel() {
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

    @Override
    public void run() {
        System.out.println("title: " + title + "url: " + url);
        try {
            getPageFromWeb(url);
            initExcel();
            getActualUrl();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
