/**
 * Created by Masachi on 2017/4/18.
 */
public class CrawlData implements Runnable {
    private String title;
    private String url;
    public CrawlData(String title, String url) {
        this.url = url;
        this.title = title;
    }

    @Override
    public void run() {
        System.out.println("title: " + title + "url: " + url);
    }
}
