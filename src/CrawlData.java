import com.gargoylesoftware.htmlunit.BrowserVersion;
import com.gargoylesoftware.htmlunit.NicelyResynchronizingAjaxController;
import com.gargoylesoftware.htmlunit.WebClient;
import com.gargoylesoftware.htmlunit.html.HtmlPage;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;

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
    
    public CrawlData(String title, String url) {
        this.url = url;
        this.title = title;
    }

    private void getPageFromWeb(String URL) throws Exception{
        client = new WebClient(BrowserVersion.CHROME);
        client.getOptions().setJavaScriptEnabled(true); //启用JS解释器，默认为true
        client.getOptions().setCssEnabled(false); //禁用css支持
//        client.getOptions().setProxyConfig(new ProxyConfig("185.10.17.134",3128));
        client.getCookieManager().setCookiesEnabled(false);
        client.getOptions().setThrowExceptionOnScriptError(false); //js运行错误时，是否抛出异常
        client.getOptions().setThrowExceptionOnFailingStatusCode(false);
        client.getOptions().setTimeout(10000); //设置连接超时时间 ，这里是10S。如果为0，则无限期等待

        client.waitForBackgroundJavaScript(600*1000);
        client.setAjaxController(new NicelyResynchronizingAjaxController());

        HtmlPage page = client.getPage(URL);
        client.waitForBackgroundJavaScript(1000*3);
        client.setJavaScriptTimeout(0);
        //System.out.println(page);
        String pageXml = page.asXml(); //以xml的形式获取响应文本
        //System.out.println(pageXml);
//        doc = Jsoup.connect(Url).get();
        originPage = pageXml;
        doc = Jsoup.parse(pageXml);
    }

    private void getActualUrl(){
        String actualUrl = MultiThread.baseUrl + doc.getElementById("frame_content").attr("src");
        //System.out.println("actualUrl: "+ actualUrl);

        int year = Integer.parseInt(title.substring(0,4));
        if(year < 2013){
            findDataOldVersion(actualUrl);
        }
        else{
            findDataNewVersion(actualUrl);
        }
    }

    private void findDataOldVersion(String urlOld){
        try {
            getPageFromWeb(urlOld);
            String totalText = originPage.replaceAll("<[^>]+>", "").replaceAll("\r\n","").replaceAll(" ","");
            System.out.println(totalText);

            Pattern peopleOld = Pattern.compile("");
            Pattern ecomonyOld = Pattern.compile("");
            Pattern agriculturalOld = Pattern.compile("");
            Pattern industryOld = Pattern.compile("");
            Pattern investmentOld = Pattern.compile("");
            Pattern tradeOld = Pattern.compile("");
            Pattern transportOld = Pattern.compile("");
            Pattern ensuranceOld = Pattern.compile("");
            Pattern educationOld = Pattern.compile("");
            Pattern cultureOld = Pattern.compile("");
            Pattern societyOLd = Pattern.compile("");
            Pattern environmentOld = Pattern.compile("");
        }
        catch (Exception e){
            e.printStackTrace();
        }
    }

    private void findDataNewVersion(String urlNew){
        try {
            getPageFromWeb(urlNew);
            String totalText = originPage.replaceAll("<[^>]+>", "").replaceAll("\r\n", "").replaceAll(" ", "").replaceAll("   |      ","").replaceAll("\n","").replaceAll("&lt;!--.*?--&gt;","");
            System.out.println(totalText);

            Pattern peopleNew = Pattern.compile("");
            Pattern ecomonyNew = Pattern.compile("");
            Pattern agriculturalNew = Pattern.compile("");
            Pattern industryNew = Pattern.compile("");
            Pattern investmentNew = Pattern.compile("");
            Pattern tradeNew = Pattern.compile("");
            Pattern transportNew = Pattern.compile("");
            Pattern ensuranceNew = Pattern.compile("");
            Pattern educationNew = Pattern.compile("");
            Pattern cultureNew = Pattern.compile("");
            Pattern societyNew = Pattern.compile("");
            Pattern environmentNew = Pattern.compile("");
        }
        catch (Exception e){
            e.printStackTrace();
        }
    }

    @Override
    public void run() {
        System.out.println("title: " + title + "url: " + url);
        try{
            getPageFromWeb(url);
            getActualUrl();
        }
        catch (Exception e){
            e.printStackTrace();
        }
    }
}
