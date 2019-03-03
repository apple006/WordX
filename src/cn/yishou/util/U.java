package cn.yishou.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.Map.Entry;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

import javax.script.ScriptEngine;

public class U {
    private static SimpleDateFormat fDate = new SimpleDateFormat("yyyy-MM-dd");
    //
    public static String formatDate(Date date) {
        if (date == null) {
            return "";
        }
        return fDate.format(date);
    }

    public static boolean isNullOrEmpty(String... str) {
        String s;
        for (int i = 0; i < str.length; i++) {
            s = str[i];
            if (s == null || s.isEmpty()) {
                return true;
            }
        }
        return false;
    }

    public static void out(String format, Object... args) {
        String s = String.format(format, args);
        System.out.println(s);
    }

    public static String toFirstLowerCase(String s) {
        if (U.isNullOrEmpty(s)) {
            return "";
        }
        String s2;
        if (s.length() > 1) {
            s2 = (s.charAt(0) + "").toLowerCase() + s.substring(1);
        } else {
            s2 = s.toLowerCase();
        }
        return s2;
    }

    public static String toFirstUpperCase(String s) {
        if (U.isNullOrEmpty(s)) {
            return "";
        }
        String s2;
        if (s.length() > 1) {
            s2 = (s.charAt(0) + "").toUpperCase() + s.substring(1);
        } else {
            s2 = s.toUpperCase();
        }
        return s2;
    }

    public static String padding0(int code, int len) {
        String result = String.format("%0" + len + "d", code);
        return result;
    }

    public static void close(Object... objects) {
        for (int i = 0; i < objects.length; i++) {
            Object obj = objects[i];
            if (obj != null) {
                Class<?> c = obj.getClass();
                try {
                    Method method = c.getMethod("close");
                    method.invoke(obj);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public static int length(String s) {
        if (U.isNullOrEmpty(s)) {
            return 0;
        }
        int count = 0;
        for (int i = 0; i < s.length(); i++) {
            char c = s.charAt(i);
            if (c >= 0 && c <= 255) {
                count++;
            } else {
                count += 2;
            }
        }
        return count;
    }

    public static String stringCut(String s, int len) {
        if (U.isNullOrEmpty(s)) {
            return "";
        }
        int count = 0;
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < s.length(); i++) {
            char c = s.charAt(i);
            int w = (c >= 0 && c <= 255) ? 1 : 2;
            count += w;
            if (count > len) {
                return sb.toString();
            }
            sb.append(c);
        }
        return sb.toString();
    }

    public static String paddingLeft(String s, String sign, int len) {
        while (s.length() < len) {
            s = sign + s;
        }
        return s;
    }

    public static String paddingRight(String s, String sign, int len) {
        while (s.length() < len) {
            s = s + sign;
        }
        return s;
    }

    // 2006-01-02 15:04:05 009
    private static Pattern regDateFormat = Pattern.compile("(20)?06|01|02|15|04|05|0*9");

    public static String parseDateFormat(String format, int number) {
        Date date = new Date();
        return parseDateFormat(format, number, date);
    }

    public static String parseDateFormat(String format, int number, Date date) {
        Calendar cal = Calendar.getInstance();
        if (date != null) {
            cal.setTime(date);
        }
        StringBuilder sb = new StringBuilder();
        int code;
        Matcher match = regDateFormat.matcher(format);
        int p = format.length();
        if (match.find()) {
            p = match.start();
        }
        for (int i = 0; i < format.length(); i++) {
            if (i == p) {
                i = match.end() - 1;
                String s = match.group();
                switch (s) {
                case "2006":
                    code = cal.get(Calendar.YEAR);
                    sb.append(code);
                    break;
                case "06":
                    code = cal.get(Calendar.YEAR);
                    sb.append(code % 100);
                    break;
                case "01":
                    code = cal.get(Calendar.MONTH) + 1;
                    sb.append(U.padding0(code, 2));
                    break;
                case "02":
                    code = cal.get(Calendar.DATE);
                    sb.append(U.padding0(code, 2));
                    break;
                case "15":
                    code = cal.get(Calendar.HOUR_OF_DAY);
                    sb.append(U.padding0(code, 2));
                    break;
                case "04":
                    code = cal.get(Calendar.MINUTE);
                    sb.append(U.padding0(code, 2));
                    break;
                case "05":
                    code = cal.get(Calendar.SECOND);
                    sb.append(U.padding0(code, 2));
                    break;
                default:
                    if (Pattern.matches("^0*9$", s)) {
                        sb.append(U.padding0(number, s.length()));
                    } else {
                        U.out("奇怪format=%s, s=%s", format, s);
                    }
                }
            } else {
                char c = format.charAt(i);
                switch (c) {
                case '\\':
                    if (i + 1 < format.length()) {
                        i++;
                        sb.append(format.charAt(i));
                    }
                    break;
                default:
                    sb.append(c);
                }
            }
            if (i >= p) {
                if (match.find()) {
                    p = match.start();
                } else {
                    p = format.length();
                }
            }
        }
        return sb.toString();
    }

    public static LinkedHashMap<String, byte[]> unzip(File file) {
        FileInputStream is;
        try {
            is = new FileInputStream(file);
        } catch (Throwable e) {
            throw new RuntimeException(e);
        }
        return unzip(is);
    }

    public static LinkedHashMap<String, byte[]> unzip(InputStream is) {
        try {
            LinkedHashMap<String, byte[]> map = new LinkedHashMap<>();
            ZipInputStream zis = new ZipInputStream(is);
            ZipEntry entry = null;
            while ((entry = zis.getNextEntry()) != null) {
                if (entry.isDirectory()) {
                    map.put(entry.getName(), null);
                    // System.out.println("Directory: " + entry.getName());
                } else {

                    byte[] data = new byte[(int) entry.getSize()];
                    int count = 0;
                    while (count < entry.getSize()) {
                        int len = zis.read(data, count, data.length - count);
                        if (len == -1) {
                            break;
                        }
                        count += len;
                    }
                    zis.closeEntry();
                    map.put(entry.getName(), data);
                }
                // System.out.println("File: " + entry.getName());
            }
            zis.close();
            is.close();
            return map;
        } catch (Throwable e) {
            throw new RuntimeException(e);
        }
    }

    public static void zip(File file, LinkedHashMap<String, byte[]> map) {
        FileOutputStream os;
        try {
            os = new FileOutputStream(file);
        } catch (Throwable e) {
            throw new RuntimeException(e);
        }
        zip(os, map);
    }

    public static void zip(OutputStream os, LinkedHashMap<String, byte[]> map) {
        try {
            ZipOutputStream zos = new ZipOutputStream(os);
            for (Entry<String, byte[]> pair : map.entrySet()) {
                ZipEntry entry = new ZipEntry(pair.getKey());
                zos.putNextEntry(entry);
                if (pair.getValue() != null) {
                    zos.write(pair.getValue());
                }
            }
            zos.close();
            os.close();
        } catch (Throwable e) {
            throw new RuntimeException(e);
        }
    }

    public static void log(String format, Object... args) {
        String s = String.format(format, args);
        System.out.println(s);
    }

    public static void logError(String format, Object... args) {
        String s = String.format(format, args);
        System.out.println(s);
    }

    public static String toJSON(Object obj) {
        if (obj == null) {
            return null;
        }
        return obj.toString();
    }
}
