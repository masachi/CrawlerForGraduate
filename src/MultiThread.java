import com.gargoylesoftware.htmlunit.BrowserVersion;
import com.gargoylesoftware.htmlunit.NicelyResynchronizingAjaxController;
import com.gargoylesoftware.htmlunit.WebClient;
import com.gargoylesoftware.htmlunit.html.HtmlPage;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.util.ArrayList;
import java.util.HashMap;

/**
 * Created by Masachi on 2017/4/18.
 */
public class MultiThread {
    public static String baseUrl = "http://www.dytj.gov.cn/";
    private static String URL = "http://www.dytj.gov.cn/article_show.aspx?typeid=9";
    private static ArrayList<String> subUrl = new ArrayList<>();
    private static ArrayList<String> title = new ArrayList<>();
    private static Document doc = null;
    private static WebClient wc = new WebClient(BrowserVersion.CHROME);

    private static void getPageFromWeb() throws Exception{
        wc.getOptions().setJavaScriptEnabled(true); //启用JS解释器，默认为true
        wc.getOptions().setCssEnabled(false); //禁用css支持
//        wc.getOptions().setProxyConfig(new ProxyConfig("185.10.17.134",3128));
        wc.getCookieManager().setCookiesEnabled(false);
        wc.getOptions().setThrowExceptionOnScriptError(false); //js运行错误时，是否抛出异常
        wc.getOptions().setThrowExceptionOnFailingStatusCode(false);
        wc.getOptions().setTimeout(10000); //设置连接超时时间 ，这里是10S。如果为0，则无限期等待

        wc.waitForBackgroundJavaScript(600*1000);
        wc.setAjaxController(new NicelyResynchronizingAjaxController());

        HtmlPage page = wc.getPage(URL);
        wc.waitForBackgroundJavaScript(1000*3);
        wc.setJavaScriptTimeout(0);
//        System.out.println(page);
        String pageXml = page.asXml(); //以xml的形式获取响应文本
//        doc = Jsoup.connect(Url).get();
        doc = Jsoup.parse(pageXml);

        wc = null;

        getSubUrl();
    }

    private static void getSubUrl(){
        Elements elements = doc.getElementsByClass("l_left").select("a");
        for(Element element : elements){
            subUrl.add(baseUrl + element.attr("href"));
            title.add(element.text());
        }
    }

    public static void main(String[] args){
        try{
            getPageFromWeb();
            for(int i=0;i<subUrl.size();i++){
                new Thread(new CrawlData(title.get(i), subUrl.get(i))).start();
            }
        }
        catch (Exception e){
            e.printStackTrace();
        }
    }
}
