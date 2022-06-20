package com.lg.project.notes.common.utils;

import com.alibaba.fastjson.JSON;
import org.apache.http.*;
import org.apache.http.client.HttpClient;
import org.apache.http.client.config.RequestConfig;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.config.Registry;
import org.apache.http.config.RegistryBuilder;
import org.apache.http.conn.ConnectionKeepAliveStrategy;
import org.apache.http.conn.socket.ConnectionSocketFactory;
import org.apache.http.conn.socket.PlainConnectionSocketFactory;
import org.apache.http.entity.ContentType;
import org.apache.http.entity.StringEntity;
import org.apache.http.entity.mime.HttpMultipartMode;
import org.apache.http.entity.mime.MultipartEntityBuilder;
import org.apache.http.impl.client.DefaultHttpRequestRetryHandler;
import org.apache.http.impl.client.DefaultServiceUnavailableRetryStrategy;
import org.apache.http.impl.client.HttpClientBuilder;
import org.apache.http.impl.conn.PoolingHttpClientConnectionManager;
import org.apache.http.message.BasicHeader;
import org.apache.http.message.BasicHeaderElementIterator;
import org.apache.http.protocol.HTTP;
import org.apache.http.protocol.HttpContext;
import org.apache.http.util.EntityUtils;
import org.springframework.web.multipart.MultipartFile;

import java.net.URI;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * HttpClient优化思路：
 *
 * 池化
 * 长连接
 * httpclient和httpget复用
 * 合理的配置参数（最大并发请求数，各种超时时间，重试次数）
 * 异步
 * 多读源码
 *
 * 一是单例的client，二是缓存的保活连接，三是更好的处理返回结果
 */
public class HttpClientUtils {

    public static String USER_AGENT = "Mozilla/5.0 (Windows NT 6.1;Win64; 64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.45 Safari/537.36";

    /**
     * 池化 提到连接缓存，很容易联想到数据库连接池。httpclient4提供了一个PoolingHttpClientConnectionManager 作为连接池。
     */
    private static PoolingHttpClientConnectionManager poolingHttpClientConnectionManager = null;

    static {
        Registry<ConnectionSocketFactory> registryBuilder = RegistryBuilder.<ConnectionSocketFactory>create()
                .register("http", new PlainConnectionSocketFactory()).build();
        poolingHttpClientConnectionManager = new PoolingHttpClientConnectionManager(registryBuilder);
        poolingHttpClientConnectionManager.setMaxTotal(20);
    }
    /**
     * 二是缓存的保活连接
     * 反复创建tcp连接的开销
     * tcp的三次握手与四次挥手两大裹脚布过程，对于高频次的请求来说，
     * 消耗实在太大。试想如果每次请求我们需要花费5ms用于协商过程，
     * 那么对于qps为100的单系统，1秒钟我们就要花500ms用于握手和挥手。
     * 又不是高级领导，我们程序员就不要搞这么大做派了，
     * 改成keep alive方式以实现连接复用！
     */
    public static ConnectionKeepAliveStrategy myStrategy = new ConnectionKeepAliveStrategy() {
        @Override
        public long getKeepAliveDuration(HttpResponse response, HttpContext context) {
            HeaderElementIterator it = new BasicHeaderElementIterator
                    (response.headerIterator(HTTP.CONN_KEEP_ALIVE));
            while (it.hasNext()) {
                HeaderElement he = it.nextElement();
                String param = he.getName();
                String value = he.getValue();
                if (value != null && param.equalsIgnoreCase
                        ("timeout")) {
                    return Long.parseLong(value) * 1000;
                }
            }
            return 60 * 1000;//如果没有约定，则默认定义时长为60s
        }
    };

    //单例的client
    private static HttpClient getClient(List<Header> header){
        RequestConfig requestConfig = RequestConfig.custom().setSocketTimeout(10000).setConnectTimeout(30000)
                .setConnectionRequestTimeout(30000).build();
        HttpClientBuilder builder = HttpClientBuilder.create().setDefaultHeaders(header).setUserAgent(USER_AGENT)
                .setKeepAliveStrategy(myStrategy)
                .setConnectionManager(poolingHttpClientConnectionManager).setConnectionManagerShared(true)
                .setRetryHandler(new DefaultHttpRequestRetryHandler(10, false))
                .setServiceUnavailableRetryStrategy(new DefaultServiceUnavailableRetryStrategy(5, 5000))
                .setDefaultRequestConfig(requestConfig);
        return builder.build();
    }

    /**
     * 对外提供 可扩展头部
     * @param extHeaderMap
     * @return
     */
    public static HttpClient getClient(Map<String,String> extHeaderMap){
        List<Header> headers = customBaseHeader();
        if(Objects.isNull(extHeaderMap) || extHeaderMap.size() ==0){
            return getClient(headers);
        }
        for(Map.Entry<String, String> entry:extHeaderMap.entrySet()){
            headers.add(new BasicHeader(entry.getKey(),entry.getValue()));
        }
        return getClient(headers);
    }

    /**
     * 获取上传的httpClient
     * @param extHeaderMap
     * @return
     */
    public static HttpClient getClientUpload(Map<String,String> extHeaderMap){
        List<Header> headers = new ArrayList<>();
        if(Objects.isNull(extHeaderMap) || extHeaderMap.size() ==0){
            return getClient(headers);
        }
        for(Map.Entry<String, String> entry :extHeaderMap.entrySet()){
            headers.add(new BasicHeader(entry.getKey(),entry.getValue()));
        }
        return getClient(headers);
    }


    public static String executeGet(HttpClient client,String uri,String charEncode) throws Exception{
        String result = null;
        HttpGet httpGet = new HttpGet(uri);
        try {
            HttpResponse response = client.execute(httpGet);
            int statusCode = response.getStatusLine().getStatusCode();
            HttpEntity entity = response.getEntity();
            if(HttpStatus.SC_OK == statusCode){
                result = EntityUtils.toString(entity, Charset.forName(charEncode));
            }else{
                EntityUtils.consumeQuietly(entity);
                throw new Exception();
            }
        }catch (Exception e){
            httpGet.abort();
            throw new Exception(e);
        }
        return result;
    }

    public static String executePost(HttpClient client,String uri,String charSet,Object param) throws Exception{
        String result = null;
        HttpPost httpPost = new HttpPost(uri);
        try {
            httpPost.setEntity(new StringEntity(JSON.toJSONString(param),Charset.forName(charSet)));
            HttpResponse response = client.execute(httpPost);
            int statusCode = response.getStatusLine().getStatusCode();
            HttpEntity entity = response.getEntity();
            if(HttpStatus.SC_OK == statusCode){
                result = EntityUtils.toString(entity, Charset.forName(charSet));
            }else{
                EntityUtils.consumeQuietly(entity);
                throw new Exception();
            }
        }catch (Exception e){
            httpPost.abort();
            throw new Exception(e);
        }
        return result;
    }

    /**
     * 附件批量上传 POST
     * @param client
     * @param uri
     * @param fileList 附件列表
     * @param charSet 编码格式
     * @return
     * @throws Exception
     */
    public static String executePostFile(HttpClient client, URI uri ,List<MultipartFile> fileList,String charSet) throws Exception{
        String result = null;
        HttpPost httpPost = new HttpPost(uri);
        try{
            MultipartEntityBuilder builder = MultipartEntityBuilder.create();
            builder.setCharset(Charset.forName(charSet));
            builder.setMode(HttpMultipartMode.BROWSER_COMPATIBLE);
            for (MultipartFile file:fileList){
                String originalFilename = file.getOriginalFilename();
                builder.addBinaryBody("attachments",file.getInputStream(), ContentType.APPLICATION_JSON,originalFilename);
            }
            httpPost.setEntity(builder.build());
            HttpResponse response = client.execute(httpPost);
            int statusCode = response.getStatusLine().getStatusCode();
            HttpEntity entity = response.getEntity();
            if(HttpStatus.SC_OK == statusCode){
                result = EntityUtils.toString(entity, Charset.forName(charSet));
            }else{
                EntityUtils.consumeQuietly(entity);
                throw new Exception();
            }
        }catch (Exception e){
            httpPost.abort();
            throw new Exception(e);
        }
        return result;
    }

    /**
     * 公有部分头部设置
     * @return
     */
    private static List<Header> customBaseHeader(){
        List<Header> headers = new ArrayList<>();
        headers.add(new BasicHeader("accept","*/*"));
        headers.add(new BasicHeader("accept-encoding","gzip,deflate,br"));
        headers.add(new BasicHeader("content-type","application/json"));
        return  headers;
    }


}
