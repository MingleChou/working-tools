package com.it.jiemin.common;

import javax.servlet.http.HttpServletRequest;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.Base64;

/**
 * @author JieminZhou
 * @version 1.0
 * @date 2019/12/16 17:27
 */
public class EncodeUtils {
    public static String getFilename(boolean flag, String filename) throws Exception {
        //处理中文文件名乱码问题
        // 如果是微软的浏览器，直接进行UTF-8编码
        if (flag) {
            filename = URLEncoder.encode(filename, "UTF-8");
            // java的编码方式和浏览器有略微的不同：对于空格，java编码后的结果是加号，
            // 而浏览器的编码结果是%20，因此将+替换成%20, 这样浏览器才能正确解析空格
            filename = filename.replace("+", "%20");
        }
        // 标准浏览器使用Base64编码
        else {
            Base64.Encoder encoder = Base64.getEncoder();
            filename = encoder.encodeToString(filename.getBytes(StandardCharsets.UTF_8));
            // =?utf-8?B?文件名?= 是告诉浏览器以Base64进行解码
            filename = "=?utf-8?B?" + filename + "?=";
        }
        return filename;//返回一个进行url编码的文件名
    }

    public static String getFilename(HttpServletRequest request, String filename) throws Exception {
        //处理中文文件名乱码问题
        // 获取浏览器类型，通过请求头中的User-Agent来判断
        String ua = request.getHeader("User-Agent");
        boolean IE_LT11 = ua.contains("MSIE"); // IE11以下版本
        boolean IE11 = ua.contains("rv:11.0) like Gecko"); // IE11
        boolean Edge = ua.contains("Edge"); // win10自带的Edge浏览器
        boolean flag = (IE_LT11 || IE11 || Edge);
        return getFilename(flag, filename);//返回一个进行url编码的文件名
    }
}
