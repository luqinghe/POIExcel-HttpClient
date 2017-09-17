package spider;

import java.io.IOException;
import java.security.KeyManagementException;
import java.security.KeyStoreException;
import java.security.NoSuchAlgorithmException;
import java.security.cert.X509Certificate;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import javax.net.ssl.SSLContext;

import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.NameValuePair;
import org.apache.http.client.HttpClient;
import org.apache.http.client.entity.UrlEncodedFormEntity;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.conn.ssl.SSLConnectionSocketFactory;
import org.apache.http.conn.ssl.SSLContextBuilder;
import org.apache.http.conn.ssl.TrustStrategy;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.message.BasicNameValuePair;
import org.apache.http.util.EntityUtils;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;
import org.junit.Test;

import excel.CreateExcel;
import excel.ReadExcel;

public class Spider {
	
	@Test
	public void testJsoup() {
		String url = "https://www.hnrsks.com/LinkPage/cjcx_list.aspx?id=3827&lx=3&zh=61304011307&zh2=41040319940205553X";
//		Map<String, String> dataMap = new HashMap<String, String>();
//		dataMap.put("TextBox1", "61304011307");
//		dataMap.put("TextBox2", "41040319940205553X");
		try {
			Document doc = Jsoup.connect(url).get();
			System.out.println(doc.toString());
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	@Test
	public void testHttpClient() {
		String url = "https://www.hnrsks.com/LinkPage/cjcx_dl.aspx?state=1&Id=3827&";
		
		HttpPost httppost = new HttpPost(url);
		HttpClient httpclient = createSSLClientDefault();
		List<NameValuePair> formparams = new ArrayList<NameValuePair>();
		formparams.add(new BasicNameValuePair("TextBox1", "61304011307")); 
		formparams.add(new BasicNameValuePair("TextBox2", "41040319940205553X")); 
		try {
			UrlEncodedFormEntity uefEntity = new UrlEncodedFormEntity(formparams, "UTF-8"); 
			httppost.setEntity(uefEntity);
			System.out.println("executing request " + httppost.getURI());
			HttpResponse response = httpclient.execute(httppost);
			System.out.println("status==" + response.getStatusLine().toString());
			HttpEntity entity = response.getEntity();
			if (entity != null)
				System.out.println("Response content: " +  EntityUtils.toString(entity,"UTF-8"));
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	@Test
	public void testHttpClientSSL() {
		String url = "https://www.hnrsks.com/LinkPage/cjcx_list.aspx?id=3827&lx=3&zh=61304011307&zh2=41040319940205553X";
		HttpGet httpGet = new HttpGet(url);
		HttpClient httpclient = createSSLClientDefault();
		try {
			HttpResponse response = httpclient.execute(httpGet);
			System.out.println("status==" + response.getStatusLine().toString());
			HttpEntity entity = response.getEntity();
			if (entity != null) {
				String content = new String(EntityUtils.toString(entity,"UTF-8"));
//				System.out.println("Response content: " +  content);
				Document document = Jsoup.parse(content);
				// 行政能力成绩
				Elements xznlcjEle = document.select("#DataGrid1 > tbody > tr:nth-child(2) > td:nth-child(3)");
				String xznlcj = xznlcjEle.text();
				System.out.println(xznlcj);
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * 创建https的client
	 * @return
	 */
	public static CloseableHttpClient createSSLClientDefault() {
		try {
			@SuppressWarnings("deprecation")
			SSLContext sslContext = new SSLContextBuilder().loadTrustMaterial(
					null, new TrustStrategy() {
						// 信任所有
						public boolean isTrusted(X509Certificate[] chain,
								String authType) {
							return true;
						}
					}).build();
			SSLConnectionSocketFactory sslsf = new SSLConnectionSocketFactory(sslContext);
			return HttpClients.custom().setSSLSocketFactory(sslsf).build();
		} catch (KeyManagementException e) {
			e.printStackTrace();
		} catch (NoSuchAlgorithmException e) {
			e.printStackTrace();
		} catch (KeyStoreException e) {
			e.printStackTrace();
		}
		return HttpClients.createDefault();
	}
	
	
	public static List<Map<String, String>> getHttpDate(List<Map<String, String>> dataList) throws Exception {
		String url = "https://www.hnrsks.com/LinkPage/cjcx_list.aspx?id=3827&lx=3";
		HttpClient httpclient = createSSLClientDefault();
		for (Map<String, String> dataMap : dataList) {
			HttpGet httpGet = new HttpGet(url + "&zh=" + dataMap.get("zkzh") + "&zh2=" + dataMap.get("sfzh"));
			HttpResponse response = httpclient.execute(httpGet);
			System.out.println("status==" + response.getStatusLine().toString());
			HttpEntity entity = response.getEntity();
			String content = new String(EntityUtils.toString(entity,"UTF-8"));
			Document document = Jsoup.parse(content);
			// 行政能力成绩
			Elements xznlcjEle = document.select("#DataGrid1 > tbody > tr:nth-child(2) > td:nth-child(1)");
			String xznlcj = xznlcjEle.text();
			dataMap.put("xznlcj", xznlcj);
			
			// 申论成绩
			Elements slcjEle = document.select("#DataGrid1 > tbody > tr:nth-child(2) > td:nth-child(2)");
			String slcj = slcjEle.text();
			dataMap.put("slcj", slcj);
			
			// 专业成绩
			Elements zycjEle = document.select("#DataGrid1 > tbody > tr:nth-child(2) > td:nth-child(3)");
			String zycj = zycjEle.text();
			dataMap.put("zycj", zycj);
			
			// 总成绩
			Elements zcjEle = document.select("#DataGrid1 > tbody > tr:nth-child(2) > td:nth-child(4)");
			String zcj = zcjEle.text();
			dataMap.put("zcj", zcj);
			
			document.clone();
		}
		return dataList;
	}
	
	public static void main(String[] args) throws Exception {
		List<Map<String, String>> dataList = ReadExcel.readWorkBook("C:\\Users\\qinghe\\Desktop\\test1.xlsx");
		Spider.getHttpDate(dataList);
		CreateExcel.ExportWorkBook(dataList, null);
	}
}