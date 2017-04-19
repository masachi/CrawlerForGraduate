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
        findData(actualUrl);
    }

    private void findData(String urlOld) {
        try {
            ArrayList<String> temp = new ArrayList<>();
            getPageFromWeb(urlOld);
            String totalText = originPage.replaceAll("<[^>]+>", "").replaceAll("\r\n", "").replaceAll(" ", "");
            //System.out.println(totalText);

            Matcher m = Pattern.compile("(\\d+\\.\\d+|\\d+)").matcher(totalText);

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


        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void initExcel(){
        wb = new XSSFWorkbook();
        initPeopleSheet();
        OutputExcel();
    }

    private void initPeopleSheet(){
        Sheet peopleSheet = wb.createSheet("综合");
        Row row = null;
        Cell cell = null;
        row = peopleSheet.createRow(0);
        row.createCell(0).setCellValue("233");
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
