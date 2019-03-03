package cn.yishou.util;

import java.io.BufferedReader;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

public class HttpUtils {
    public static class SimpleResponse {
        public byte[] data;
        public Map<String, List<String>> header;

        public String getHead(String name) {
            if (header != null && header.containsKey(name)) {
                return header.get(name).get(0);
            }
            return null;
        }
    }

    public static String doGet(String url, HashMap<String, String> headerMap) {
        SimpleResponse rsp = doGetBase(url, headerMap);
        if (rsp.data != null) {
            String result = new String(rsp.data);
            return result;
        }
        return "";
    }

    /**
     * 向指定URL发送GET方法的请求
     * 
     * @param url
     *            发送请求的URL
     * @param param
     *            请求参数，请求参数应该是 name1=value1&name2=value2 的形式。
     * @return URL 所代表远程资源的响应结果
     */
    public static SimpleResponse doGetBase(String url, HashMap<String, String> headerMap) {
        SimpleResponse result = new SimpleResponse();
        byte[] data = null;
        InputStream is = null;
        try {
            String urlNameString = url;
            URL realUrl = new URL(urlNameString);
            // 打开和URL之间的连接
            HttpURLConnection connection = (HttpURLConnection) realUrl.openConnection();
            // 设置通用的请求属性
            connection.setRequestProperty("accept", "*/*");
            connection.setRequestProperty("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1;SV1)");
            if (headerMap != null) {
                for (Entry<String, String> e : headerMap.entrySet()) {
                    connection.setRequestProperty(e.getKey(), e.getValue());
                }
            }
            // 建立实际的连接
            connection.connect();
            // 获取所有响应头字段
            Map<String, List<String>> map = connection.getHeaderFields();
            result.header = map;
            // 遍历所有的响应头字段
            if (connection.getResponseCode() != 200) {
                for (String key : map.keySet()) {
                    System.out.println(key + "--->" + map.get(key));
                }
            }
            if (connection.getResponseCode() == 200) {
                if (map.containsKey("Content-Length")) {
                    int len = Integer.parseInt(map.get("Content-Length").get(0));
                    data = new byte[len];
                    int offset = 0;
                    is = connection.getInputStream();
                    while (true) {
                        int len2 = is.read(data, offset, len - offset);
                        if (len2 == -1 || offset + len2 >= len) {
                            break;
                        }
                        offset += len2;
                    }
                    result.data = data;
                } else if ("chunked".equals(map.get("Transfer-Encoding").get(0))) {
                    data = chunked(connection.getInputStream());
                    result.data = data;
                }
            }
        } catch (Exception e) {
            U.out("访问失败:%s", url);
            throw new RuntimeException(e);
        } finally {
            try {
                if (is != null) {
                    is.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return result;
    }

    /**
     * 向指定 URL 发送POST方法的请求
     * 
     * @param url
     *            发送请求的 URL
     * @param param
     *            请求参数，请求参数应该是 name1=value1&name2=value2 的形式。
     * @return 所代表远程资源的响应结果
     */
    public static String doPost(String url, HashMap<String, String> header, String param) {
        PrintWriter out = null;
        BufferedReader in = null;
        String result = "";
        try {
            URL realUrl = new URL(url);
            // 打开和URL之间的连接
            HttpURLConnection conn = (HttpURLConnection) realUrl.openConnection();
            // 设置通用的请求属性
            conn.setRequestProperty("accept", "*/*");
            conn.setRequestProperty("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1;SV1)");
            if (header != null && header.size() > 0) {
                for (Entry<String, String> entry : header.entrySet()) {
                    conn.setRequestProperty(entry.getKey(), entry.getValue());
                }
            }
            // 发送POST请求必须设置如下两行
            conn.setDoOutput(true);
            conn.setDoInput(true);
            // 获取URLConnection对象对应的输出流
            out = new PrintWriter(conn.getOutputStream());
            // 发送请求参数
            out.print(param);
            // flush输出流的缓冲
            out.flush();
            //
            if (conn.getResponseCode() != 200) {
                U.out("响应code=%d", conn.getResponseCode());
                return null;
            }
            // 定义BufferedReader输入流来读取URL的响应
            in = new BufferedReader(new InputStreamReader(conn.getInputStream()));
            String line;
            while ((line = in.readLine()) != null) {
                result += line;
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        // 使用finally块来关闭输出流、输入流
        finally {
            try {
                if (out != null) {
                    out.close();
                }
                if (in != null) {
                    in.close();
                }
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        }
        return result;
    }

    public static byte[] chunked(InputStream in) throws Exception {
        ByteArrayOutputStream tmpos = new ByteArrayOutputStream(4);
        ByteArrayOutputStream bytes = new ByteArrayOutputStream();
        int data = -1;
        int[] aaa = new int[2];
        byte[] aa = null;

        while ((data = in.read()) >= 0) {
            aaa[0] = aaa[1];
            aaa[1] = data;
            if (aaa[0] == 13 && aaa[1] == 10) {
                aa = tmpos.toByteArray();
                int num = 0;
                try {
                    num = Integer.parseInt(new String(aa, 0, aa.length - 1)
                        .trim(), 16);
                } catch (Exception e) {
                    System.out.println("aa.length:" + aa.length);
                    e.printStackTrace();
                }

                if (num == 0) {

                    in.read();
                    in.read();
                    return bytes.toByteArray();
                }
                aa = new byte[num];
                int sj = 0, ydlen = num, ksind = 0;
                while ((sj = (in.read(aa, ksind, ydlen))) < ydlen) {
                    ydlen -= sj;
                    ksind += sj;
                }

                bytes.write(aa);
                in.read();
                in.read();
                tmpos.reset();
            } else {
                tmpos.write(data);
            }
        }
        return tmpos.toByteArray();
    }
}