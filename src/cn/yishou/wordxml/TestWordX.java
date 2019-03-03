package cn.yishou.wordxml;

import java.io.File;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.StandardOpenOption;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;

import cn.yishou.util.U;

public class TestWordX {
    public static class Entity {
        public int sex;
    }

    public static void main(String[] args) {
        test();
    }

    public static void test() {
        String inPath = "test.docx";
        String outPath = "E:\\WordXTest.docx";
        // 准备数据
        Entity entity = new Entity();
        entity.sex = 2;
        //
        HashMap<String, String> mapData = new HashMap<>();
        ArrayList<HashMap<String, String>> list = new ArrayList<>();
        for (int i = 0; i < 6; i++) {
            mapData = new HashMap<>();
            mapData.put("RealName", "小何" + i);
            mapData.put("Phone", "Phone" + i);
            mapData.put("IdCard", "IdCard" + i);
            list.add(mapData);
        }
        //
        HashMap<String, Object> dataMap = new HashMap<>();
        dataMap.put("SexEnum", "0最美 1少男 2少女");
        dataMap.put("SignUrl0", "https://upload.wikimedia.org/wikipedia/zh/b/b1/Avon-Tyres-Logo.png");
        dataMap.put("SignUrl1", "http://eq.10jqka.com.cn/logo/LAMR.png");
        dataMap.put("Now", new Date());
        dataMap.put("List", list);
        dataMap.put("Pi", Math.PI * 1000);
        // 准备文件
        try {
            // 打开Word文件
            InputStream is = TestWordX.class.getClassLoader().getResourceAsStream(inPath);
            // 绑定数据 真正调用只此一句
            byte[] data = WordX.parse(is, entity, dataMap);
            // 写文件
            File f2 = new File(outPath);
            Files.write(f2.toPath(), data, StandardOpenOption.CREATE);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        //
        U.out("outPath=%s", outPath);
    }

}
