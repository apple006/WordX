package cn.yishou.wordxml;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.w3c.dom.ls.DOMImplementationLS;
import org.w3c.dom.ls.LSSerializer;

import cn.yishou.util.HttpUtils;
import cn.yishou.util.HttpUtils.SimpleResponse;
import cn.yishou.util.U;

public class WordX {
    private static int debug = 0;
    // 等宽填充
    private static Pattern regEW = Pattern.compile("_EW\\b");
    // 最大宽
    private static Pattern regW = Pattern.compile("_(\\d+)W\\b");
    // 单选框
    private static Pattern regRadio = Pattern.compile("_(\\d+)_Radio$");
    // 复选框
    private static Pattern regCheck = Pattern.compile("_Check$");
    // 格式化
    private static Pattern regFmt = Pattern.compile("_Fmt$");
    // 整型变枚举
    private static Pattern regEnum = Pattern.compile("_Enum$");
    // List循环 min最少要循环的次数 max最多只能循环的次数 offset从第几行开始(首行为0)
    private static Pattern regList = Pattern.compile("_(\\d*)List(\\d*)_(\\d*)(Copy|Fill)_([^_]+)$");
    // 图片
    private static Pattern regImg = Pattern.compile("_Img$");
    // 删除
    private static Pattern regDelete = Pattern.compile("^X_Delete(_Last)?$");
    // 解释枚举字符串
    private static Pattern regParseEnum = Pattern.compile("^\\d+");

    private static class ListInfo {
        public List<?> listData;
        public int index;
        public String type;
        public String itemName;
        public Object item;
        public List<Node> listNode;
        public HashMap<Node, Node> bookmarkMap;
        public int min;
        public int max;
        public int offset;
        public int count = -1;

        public Object next() {
            if (index >= listData.size()) {
                index++;
                item = null;
                return null;
            }
            item = listData.get(index++);
            out("%s 循环 index=%d", itemName, index);
            return item;
        }

        public boolean isEnd() {
            if (listData == null) {
                return true;
            }
            if (count == -1) {
                count = Math.max(min + offset, listData.size());
                if (max > 0) {

                    count = Math.min(max + offset, count);
                }
            }
            if (index > count) {
                out("index%d>count%d", index, count);
            }
            return index > count;
        }
    }

    private static class Context {
        LinkedHashMap<String, byte[]> fileMap;
        Document doc;
        Document docRel;
        Object entity;
        Object obj;
        int relationMaxId = 0;
        Map<String, Element> relationMap;
        Map<String, Integer> relationUseMap;
        Map<String, Map<Integer, String>> enumsMap;
        Map<String, Object> dataMap;
        ListInfo listInfo;
    }

    public static byte[] parse(File f, Object entity, Map<String, Object> dataMap) {
        FileInputStream is;
        try {
            is = new FileInputStream(f);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
        return parse(is, entity, dataMap);
    }

    public static byte[] parse(InputStream is, Object entity, Map<String, Object> dataMap) {
        LinkedHashMap<String, byte[]> fileMap = U.unzip(is);
        byte[] data = fileMap.get("word/document.xml");
        ByteArrayInputStream bis = new ByteArrayInputStream(data);
        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document doc = builder.parse(bis);
            Node root = doc.getElementsByTagName("w:document").item(0);
            //
            Context context = new Context();
            context.fileMap = fileMap;
            context.doc = doc;
            context.dataMap = dataMap;
            context.entity = entity;
            context.enumsMap = new HashMap<>();
            context.relationUseMap = new HashMap<>();
            getRelationships(context);
            //
            parseOne(root, context);
            //
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            Transformer former = TransformerFactory.newInstance().newTransformer();
            former.transform(new DOMSource(doc), new StreamResult(bos));
            U.close(is, bos);
            fileMap.put("word/document.xml", bos.toByteArray());
            //
            if (context.relationMap.size() > 0) {
                bos = new ByteArrayOutputStream();
                former = TransformerFactory.newInstance().newTransformer();
                former.transform(new DOMSource(context.docRel), new StreamResult(bos));
                U.close(is, bos);
                fileMap.put("word/_rels/document.xml.rels", bos.toByteArray());
            }
            //
            bos = new ByteArrayOutputStream();
            U.zip(bos, fileMap);
            return bos.toByteArray();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private static void getRelationships(Context context) {
        byte[] data = context.fileMap.get("word/_rels/document.xml.rels");
        ByteArrayInputStream bis = new ByteArrayInputStream(data);
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        DocumentBuilder builder;
        Document docRel;
        try {
            builder = factory.newDocumentBuilder();
            docRel = builder.parse(bis);
        } catch (Throwable e) {
            throw new RuntimeException(e);
        }
        context.relationMap = new HashMap<>();
        NodeList ships = docRel.getElementsByTagName("Relationship");
        for (int i = 0; i < ships.getLength(); i++) {
            Element ship = (Element) ships.item(i);
            String id = ship.getAttribute("Id");
            int intId = Integer.parseInt(id.substring(3));
            if (intId > context.relationMaxId) {
                context.relationMaxId = intId;
            }
            if (ship.getAttribute("Type").endsWith("/image")) {
                context.relationMap.put(id, ship);
            }
        }
        U.close(bis);
        context.docRel = docRel;
    }

    private static void parseOne(Node root, Context context) {
        try {
            Document doc = root.getOwnerDocument();
            Object[] box = getBookmark(root);
            ArrayList<Node> listAll = (ArrayList<Node>) box[0];
            HashMap<Node, Node> bookmarkMap = (HashMap<Node, Node>) box[1];
            out("listAll.size=%d", listAll.size());
            for (int k = 0; k < listAll.size(); k++) {
                Node start = listAll.get(k);
                Node end = bookmarkMap.get(start);

                String tagName = start.getNodeName();
                if (!"w:bookmarkStart".equals(tagName)) {
                    continue;
                }
                String startName = getAttribute(start, "w:name");
                if ("_GoBack".equals(startName)) {
                    continue;
                }
                if (end == null) {
                    out("没有匹配的End start=%s", nodeString(start));
                    continue;
                }
                out("解释标记 %s", startName);
                String startId = getAttribute(start, "w:id");
                String endId = getAttribute(end, "w:id");
                //
                startName = startName.replaceAll("_\\d+$", "");
                String value;
                //
                context.obj = context.entity;
                if (context.listInfo != null) {
                    if (context.listInfo.index == 3 && startName.equals("ShipName_EW")) {
                        startName = startName + "";
                    }
                    if (startName.startsWith(context.listInfo.itemName)) {
                        context.obj = context.listInfo.item;
                        startName = startName.substring(context.listInfo.itemName.length());
                    }
                }
                if (context.listInfo == null) {
                    Matcher match = regList.matcher(startName);
                    if (match.find()) {
                        moveNodeForLoop(listAll, start, end);
                        //
                        context.listInfo = new ListInfo();
                        context.listInfo.type = match.group(4);
                        context.listInfo.itemName = match.group(5) + "_";
                        String str = match.group(1);
                        if (!str.isEmpty()) {
                            context.listInfo.min = Integer.parseInt(str);
                        }
                        str = match.group(2);
                        if (!str.isEmpty()) {
                            context.listInfo.max = Integer.parseInt(str);
                        }
                        str = match.group(3);
                        if (!str.isEmpty()) {
                            context.listInfo.offset = Integer.parseInt(str);
                        }
                        startName = match.replaceAll("");
                        out("循环 listInfo=%s", U.toJSON(context.listInfo));
                        context.listInfo.listData = (List<?>) getValue(context, startName, null);
                        box = findBookmarkRange(start, end);
                        context.listInfo.listNode = (List<Node>) box[0];
                        int lastIndex = (int) box[1];
                        start.getParentNode().removeChild(start);
                        end.getParentNode().removeChild(end);
                        if (context.listInfo.listData == null) {
                            System.out.println("list==null " + startName);
                        } else {
                            Node lastNode = context.listInfo.listNode.get(lastIndex);
                            Node ref = lastNode.getNextSibling();
                            if (ref == null) {
                                out("lastNode=%s", nodeString(lastNode));
                            }
                            if (ref != null && ref.getNodeName().startsWith("w:bookmark")) {
                                ref = context.listInfo.listNode.get(context.listInfo.listNode.size() - 1)
                                    .getNextSibling();
                            }
                            out("开始之前listNode.size=%d; ref=%s; lastNode=%s", context.listInfo.listNode.size(),
                                nodeString(ref), nodeString(lastNode));
                            Node parentNode = lastNode.getParentNode();
                            //
                            Node p0 = start.getOwnerDocument().createElement("w:p"); // 制造假的循环根元素
                            for (Node n : context.listInfo.listNode) {
                                if (n.getParentNode() == null) {
                                    continue;
                                }
                                Node n2 = n.getParentNode().removeChild(n);
                                p0.appendChild(n2);
                            }
                            //
                            int ii = 0;
                            context.listInfo.index = context.listInfo.offset;
                            while (true) {
                                ii++;
                                context.listInfo.next();
                                if (context.listInfo.isEnd()) {
                                    break;
                                }
                                out("listInfo.item=%s", U.toJSON(context.listInfo.item));
                                Node root2 = p0.cloneNode(true);
                                startName = getAttribute(start, "w:name");
                                startId = getAttribute(start, "w:id");
                                out("循环 %d 开始 tagName=%s; startName=%s,startId=%s", ii, start.getNodeName(), startName,
                                    startId);
                                parseOne(root2, context);
                                out("循环 %d 结束 tagName=%s; startName=%s,startId=%s", ii, end.getNodeName(), startName,
                                    startId);
                                if (ref != null && parentNode != ref.getParentNode()) {
                                    out("父=%s; 引父=%s; 引=%s", nodeString(parentNode), nodeString(ref.getParentNode()),
                                        nodeString(ref));
                                }
                                while (root2.hasChildNodes()) {
                                    Node n = root2.removeChild(root2.getFirstChild());
                                    if (n.getNodeName().startsWith("w:bookmark")) {
                                        continue;
                                    }
                                    parentNode.insertBefore(n, ref);
                                }
                            }
                            p0 = null;
                            //
                            // 跳到list end处
                            while (start != end) {
                                k++;
                                start = listAll.get(k);
                            }
                            out("成功跳到list end处");
                        }
                        context.listInfo = null;
                        continue;
                    }
                }
                //
                boolean nobody = startName.startsWith("X_Delete");
                ArrayList<Node> nodes = findNodes(start, end, "w:r", true, true);
                if (!nobody && nodes.size() == 0) {
                    U.log("name=%s; not find w:r", startName);
                    continue;
                }
                ArrayList<Node> nodes2 = findNodes(start, end, "w:t", true, true);
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < nodes.size(); i++) {
                    Element n = (Element) nodes.get(i);
                    sb.append(n.getTextContent());
                    if (i > 0) {
                        n.getParentNode().removeChild(n);
                    }
                }
                String format = sb.toString();
                Element t = null;
                if (nodes.size() > 0) {
                    t = (Element) nodes.get(0).getLastChild();
                    if (nodes.size() > 1 && nodes2.size() > 0) {
                        if (!"w:t".equals(t.getNodeName())) {
                            t = (Element) nodes2.get(0);
                        }
                        t.setTextContent(sb.toString());
                    }
                    t.setAttribute("xml:space", "preserve");
                }

                value = null;
                //
                Matcher match = regEW.matcher(startName);
                boolean isEW = false;
                boolean isSetValue = false;
                if (match.find()) {
                    startName = match.replaceAll("");
                    isEW = true;
                }
                match = regW.matcher(startName);
                int maxLen = 0;
                if (match.find()) {
                    maxLen = Integer.parseInt(match.group(1));
                    startName = match.replaceAll("");
                }
                match = regRadio.matcher(startName); // □ F0A3 F052 163 82
                if (match.find()) {
                    startName = startName.substring(0, match.start());
                    value = getStringValue(context, startName, format);
                    String value2 = match.group(1);
                    if ("00".equals(value2)) {
                        if (format.indexOf(value) != -1) {
                            Element r = (Element) t.getParentNode();
                            if (format.indexOf("□") != -1) {
                                Element run = (Element) r.cloneNode(true);
                                Node t2 = run.getElementsByTagName("w:t").item(0);
                                Element sym = createCheckedBox(doc);
                                run.replaceChild(sym, t2);
                                Node parent = r.getParentNode();
                                if (format.trim().startsWith("□")) {
                                    parent.insertBefore(run, r);
                                } else {
                                    parent.insertBefore(run, r.getNextSibling());
                                }
                                //
                                value = format.replace("□", "");
                                t.setTextContent(value);
                                out("有点怪preserve:%s", nodeString(t));
                            } else {
                                NodeList arr = r.getElementsByTagName("w:sym");
                                if (arr.getLength() > 0) {
                                    Element t2 = (Element) arr.item(0);
                                    t2.setAttribute("w:char", "F052");
                                }
                            }
                        }
                    } else if (value.equals(value2)) {
                        Node r = t.getParentNode();
                        Element sym = createCheckedBox(doc);
                        r.replaceChild(sym, t);
                    }
                    continue;
                }
                match = regCheck.matcher(startName);
                if (match.find()) {
                    startName = match.replaceAll("");
                    value = getStringValue(context, startName, format);
                    if (!"0".equals(value) && !U.isNullOrEmpty(value)) {
                        Node r = t.getParentNode();
                        Element sym = createCheckedBox(doc);
                        r.replaceChild(sym, t);
                    }
                    continue;
                }
                match = regEnum.matcher(startName);
                if (match.find()) {
                    startName = startName.substring(0, match.start());
                    Integer intValue = (Integer) getValue(context, startName, format);
                    value = "";
                    if (intValue != null) {
                        value = getEnumItem(context, startName, intValue);
                    }
                    isSetValue = true;
                }
                match = regImg.matcher(startName);
                if (match.find()) {
                    processImg(context, t, match, startName, format);
                    continue;
                }
                match = regDelete.matcher(startName);
                if (match.find()) {
                    if (U.isNullOrEmpty(match.group(1))) {
                        box = findBookmarkRange(start, end);
                        ArrayList<Node> list = (ArrayList<Node>) box[0];
                        int lastIndex = (Integer) box[1];
                        out("删除 list.size=%s", list.size());
                        for (Node n : list) {
                            out("删除 %s", n.getTextContent());
                            n.getParentNode().removeChild(n);
                        }
                    } else if (context.listInfo != null && context.listInfo.index == context.listInfo.count) {
                        Node r = t.getParentNode();
                        r.getParentNode().removeChild(r);

                    }
                    continue;
                }
                // 普通字段
                if (!isSetValue) {
                    value = getStringValue(context, startName, format);
                }
                if (value != null) {
                    String oldValue = t.getTextContent();
                    if (maxLen > 0) {
                        value = U.stringCut(value, maxLen);
                    }
                    if (isEW) {
                        int count = U.length(oldValue) - U.length(value);
                        if (count > 0) {
                            int left = count / 2;
                            int rigth = count - left;
                            value = U.paddingLeft("", " ", left) + value + U.paddingRight("", " ", rigth);
                        }
                    }
                    out("　　　　 %s=%s; old=%s; tagName=%s", startName, value, oldValue,
                        t.getNodeName());
                    t.setTextContent(value);
                } else {
                    out("　　　　 %s没有value; t=[%s]", startName, t.getTextContent());
                }
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private static Object[] getBookmark(Node root) {
        HashMap<Node, Node> map = new HashMap<Node, Node>();
        HashMap<String, Node> map2 = new HashMap<String, Node>();
        ArrayList<Node> list = findNodes(
            root,
            null,
            "w:bookmarkStart;w:bookmarkEnd",
            true, false);
        out("bookmarkCount=%d", list.size());
        for (int i = 0; i < list.size(); i++) {
            Element node = (Element) list.get(i);
            String id = node.getAttribute("w:id");
            String tagName = node.getNodeName();
            if ("w:bookmarkStart".equals(tagName)) {
                boolean nearLoop = false;// 是否紧挨着循环标记
                Element n = node;
                while (true) {
                    n = (Element) n.getNextSibling();
                    if (n == null || !n.getTagName().startsWith("w:bookmark")) {
                        break;
                    }
                    String name = n.getAttribute("w:name");
                    if (regList.matcher(name).find()) {
                        nearLoop = true;
                        break;
                    }
                }
                if (nearLoop) {
                    list.remove(i);
                    i--;
                    continue;
                }
                map2.put(id, node);
                map.put(node, null);
            } else {
                Node start = map2.get(id);
                map.put(start, node);
                map.put(node, start);
            }
        }
        return new Object[] { list, map };
    }

    private static String getStringValue(
        Context context,
        String key,
        String format) {
        Object value = getValue(context, key, format);
        if (value == null) {
            if (U.isNullOrEmpty(format)) {
                return "";
            }
            return format.replaceAll(".", " ");
        }
        return value.toString();
    }

    private static Object getValue(
        Context context,
        String key,
        String format) {
        Matcher match = regFmt.matcher(key);
        if (match.find()) {
            key = key.substring(0, match.start());
            out("格式化: %s; format=%s", key, format);
        } else {
            format = null;
        }
        Object value = null;
        if (context.listInfo != null && "ListIndex".equals(key)) {
            if (context.listInfo.index > context.listInfo.listData.size()) {
                return "";
            }
            value = context.listInfo.index;
        }
        if (value == null && context.dataMap != null) {
            value = context.dataMap.get(key);
        }
        if (value == null && context.obj != null) {
            if (context.obj instanceof Map) {
                Map<String, Object> map = (Map<String, Object>) context.obj;
                value = map.get(key);
            } else {
                Class<?> clz = context.obj.getClass();
                Field f = null;
                try {
                    f = clz.getField(key);
                    value = f.get(context.obj);
                } catch (Exception e) {
                    try {
                        String key2 = U.toFirstLowerCase(key);
                        f = clz.getField(key2);
                        value = f.get(context.obj);
                    } catch (Exception e1) {
                        try {
                            String key3 = U.toFirstUpperCase(key);
                            Method method = clz.getMethod("get" + key3);
                            value = method.invoke(context.obj);
                        } catch (Exception e2) {
                        }
                    }
                }
            }
        }
        if (value != null) {
            if (value instanceof Date) {
                if (!U.isNullOrEmpty(format)) {
                    String s = U.parseDateFormat(format, 0, (Date) value);
                    return s;
                } else {
                    return U.formatDate((Date) value);
                }
            } else if (format != null) {
                if (!format.contains("%")) {
                    format = "%" + format;
                }
                value = String.format(format, value);
            }
            return value;
        } else if (format != null) {
            return format.replaceAll(".", " ");
        }
        return null;
    }

    private static ArrayList<Node> findNodes(Node start, Node end, String filter, boolean isTree, boolean findBrother) {
        ArrayList<HashMap<String, String>> list = new ArrayList<>();
        String[] arr1 = filter.split(";|,");
        for (String item1 : arr1) {
            HashMap<String, String> map = new HashMap<>();
            String[] arr2 = item1.split(" ");
            for (String item2 : arr2) {
                String[] arr3 = item2.split("=");
                if (arr3.length == 2) {
                    map.put(arr3[0], arr3[1]);
                } else if (!arr3[0].isEmpty()) {
                    map.put("tagName", arr3[0]);
                }
            }
            list.add(map);
        }
        boolean isAnd = filter.indexOf(",") > 0;
        return findNodes(start, end, list, isTree, findBrother, isAnd).list;
    }

    private static class FindResult {
        ArrayList<Node> list = new ArrayList<>();;
        boolean isEnd = false;
    }

    private static FindResult findNodes(Node start, Node end, ArrayList<HashMap<String, String>> filter,
        boolean isTree, boolean findBrother, boolean isAnd) {
        FindResult result = new FindResult();
        if (start == null) {
            return result;
        }
        if (start == end) {
            result.isEnd = true;
            return result;
        }
        //
        Node node = start;
        boolean isFirst = true;
        while (true) {
            if (!isFirst) {
                node = node.getNextSibling();
            }
            isFirst = false;
            if (node == null) {
                break;
            }
            // out("NodeName=%s", node.getNodeName());
            if (node == end) {
                result.isEnd = true;
                return result;
            }
            boolean match = isAnd;
            for (HashMap<String, String> map : filter) {
                boolean match2 = true;
                for (Entry<String, String> e : map.entrySet()) {
                    if ("tagName".equals(e.getKey())) {
                        String tagName = node.getNodeName();
                        if (!tagName.equals(e.getValue())) {
                            match2 = false;
                            break;
                        }
                    } else {
                        String attr = getAttribute(node, e.getKey());// node.getAttributes().getNamedItem(e.getKey());
                        if (attr == null || !attr.equals(e.getValue())) {
                            match2 = false;
                            break;
                        }
                    }
                }
                if (!match2 && isAnd) {
                    match = false;
                    break;
                }
                if (match2 && !isAnd) {
                    match = true;
                    break;
                }
            }
            if (match) {
                result.list.add(node);
            }
            if (isTree && node.hasChildNodes()) {
                FindResult r2 = findNodes(node.getFirstChild(), end, filter, isTree, true, isAnd);
                result.list.addAll(r2.list);
                result.isEnd = r2.isEnd;
                if (result.isEnd) {
                    return result;
                }
            }
            if (!findBrother) {
                break;
            }
        }
        //
        return result;
    }

    private static String getAttribute(Node node, String name) {
        if (node != null) {
            NamedNodeMap attrs = node.getAttributes();
            if (attrs != null) {
                Node attr = attrs.getNamedItem(name);
                if (attr != null) {
                    return attr.getNodeValue();
                }
            }
        }
        return null;
    }

    private static Object[] findBookmarkRange(Node start, Node end) {
        if (end == null) {
            out("New找范围 none end. start=%s", getAttribute(start, "w:name"));
        }
        out("找范围 start=%s; end=%s", nodeString(start), nodeString(end));
        // 附近的End标记
        ArrayList<Node> listOtherEnd = new ArrayList<>();
        Node n = end;
        while (true) {
            n = n.getPreviousSibling();
            if (n == null || !n.getNodeName().startsWith("w:bookmarkEnd")) {
                break;
            }
            listOtherEnd.add(n);
        }
        // 从End开始往上找Start
        boolean isClear = false;// 结果集是否清空过
        Element eStart = (Element) start;
        Element eEnd = (Element) end;
        String startId = getAttribute(start, "w:id");
        String filter = "tagName=w:bookmarkStart,w:id=" + startId;
        ArrayList<Node> result = new ArrayList<>();
        n = end;
        Node pre;
        boolean addParent = false;
        while (true) {
            out("加加n=%s。", n.getTextContent());
            result.add(n);
            // out("查找结果添加:%s", nodeString(n));
            ArrayList<Node> list = findNodes(n, null, filter, true, false);
            if (list.size() > 0) {
                // out("有找到啊。size=%d; node=%s", list.size(), nodeString(list.get(0)));
                if (result.size() == 1) {
                    Node n2 = n;
                    while (true) {
                        pre = n2;
                        n2 = n2.getLastChild();
                        if (n2 != null) {
                            list = findNodes(n2, null, filter, true, false);
                        }
                        if (n2 == null || list.size() == 0) {
                            result.remove(n);
                            result.add(pre);
                            break;
                        }
                    }
                }
                break;
            }
            pre = n;
            n = n.getPreviousSibling();
            while (n == null) {
                out("查找结果清空");
                for (int i = 0; i < result.size() && !addParent; i++) {
                    if (result.get(i).getNodeName().equals("w:r")) {
                        addParent = true;
                        break;
                    }
                }
                result.clear();
                isClear = true;
                pre = pre.getParentNode();
                if (addParent) {
                    result.add(pre);
                    isClear = false;
                }
                n = pre.getPreviousSibling();
            }
        }
        result.remove(end);
        result.remove(start);
        // Start是否在结果的第一个子元素中，如果不在，减少范围。
        if (result.size() == 1) {
            n = result.get(0);
            ArrayList<Node> list = findNodes(n, start.getParentNode(), "w:p", true, false);
            if (list.size() > 0) {
                result.clear();
                isClear = true;
                n = n.getLastChild();
                while (true) {
                    list = findNodes(n, null, filter, true, false);
                    result.add(n);
                    n = n.getPreviousSibling();
                    if (n == null || list.size() > 0) {
                        break;
                    }
                }
            }
        }
        //
        Collections.reverse(result);
        int lastIndex = result.size() - 1; // 最后一个有效标记位置。因为后面会追加许多End标记
        if (isClear && listOtherEnd.size() > 0) {
            for (int i = 0; i < listOtherEnd.size(); i++) {
                // out("附加标记 %s", nodeString(listOtherEnd.get(i)));
            }
            Collections.reverse(listOtherEnd);
            result.addAll(listOtherEnd);
        }
        for (int i = 0; i < result.size(); i++) {
            out("New 查找范围结果 %d：%s", i, result.get(i).getTextContent());
        }
        // out("New lastIndex= %d", lastIndex);
        //
        return new Object[] { result, lastIndex };
    }

    private static String nodeString(Node node) {
        if (node == null) {
            return "null";
        }
        DOMImplementationLS lsImpl = (DOMImplementationLS) node.getOwnerDocument().getImplementation().getFeature("LS",
            "3.0");
        LSSerializer lsSerializer = lsImpl.createLSSerializer();
        String xml = lsSerializer.writeToString(node);
        int p = xml.indexOf("\n");
        if (p > 0) {
            xml = xml.substring(p + 1);
        }
        return xml;
    }

    private static String getEnumItem(Context context,
        String type, int key) {
        Map<Integer, String> map = context.enumsMap.get(type);
        String value = null;
        if (map != null) {
            value = map.get(key);
        } else {
            String s = (String) context.dataMap.get(type + "Enum");
            if (s != null) {
                map = stringToEnum(s);
                context.enumsMap.put(type, map);
                value = map.get(key);
            }
        }
        if (value == null) {
            value = key + "";
        }
        return value;
    }

    private static Map<Integer, String> stringToEnum(String s) {
        String[] arr = s.split(" ");
        Map<Integer, String> map = new HashMap<>();
        for (String item : arr) {
            item = item.trim();
            Matcher match = regParseEnum.matcher(item);
            if (match.find()) {
                int key = Integer.parseInt(match.group());
                String value = item.substring(match.end());
                map.put(key, value);
            }
        }
        return map;
    }

    private static Element createCheckedBox(Document doc) {
        Element sym = doc.createElement("w:sym");
        sym.setAttribute("w:font", "Wingdings 2");
        sym.setAttribute("w:char", "F052");
        return sym;
    }

    private static void moveNodeForLoop(ArrayList<Node> listAll, Node start, Node end) {
        Node parent = start.getParentNode();
        Node temp;
        // 循环开始标记后的w:bookmarkEnd向上移出循环
        // .....可以不处理
        // 循环开始记前的w:bookmarkStart向下移入循环
        int startIndex = listAll.indexOf(start);
        Node n = start;
        while (true) {
            n = n.getPreviousSibling();
            if (n == null || !n.getNodeName().startsWith("w:bookmark")) {
                break;
            }
            if (n.getNodeName().equals("w:bookmarkStart")) {
                listAll.remove(n);
                startIndex--;
                listAll.add(startIndex + 1, n);
                temp = n.getNextSibling();
                n = parent.removeChild(n);
                parent.insertBefore(n, start.getNextSibling());
                n = temp;
            }
        }
        // 循环结束标记后的w:bookmarkEnd向上移入循环
        parent = end.getParentNode();
        //
        int endIndex = listAll.indexOf(end);
        n = end;
        while (true) {
            n = n.getNextSibling();
            if (n == null || !n.getNodeName().startsWith("w:bookmark")) {
                break;
            }
            if (n.getNodeName().equals("w:bookmarkEnd")) {
                listAll.remove(n);
                listAll.add(endIndex, n);
                endIndex++;
                temp = n.getPreviousSibling();
                n = parent.removeChild(n);
                out("移动n=%s; end=%s", nodeString(n), nodeString(end));
                parent.insertBefore(n, end);
                n = temp;
            }
        }
        // 循环结束标记前的w:bookmarkStart向下移出循环
        endIndex = listAll.indexOf(end);
        n = end;
        while (true) {
            n = n.getPreviousSibling();
            if (n == null || !n.getNodeName().startsWith("w:bookmark")) {
                break;
            }
            if (n.getNodeName().equals("w:bookmarkStart")) {
                listAll.remove(n);
                endIndex--;
                listAll.add(endIndex + 1, n);
                temp = n.getNextSibling();
                n = parent.removeChild(n);
                parent.insertBefore(n, end.getNextSibling());
                n = temp;
            }
        }
    }

    private static void out(String format, Object... args) {
        if (debug == 1) {
            U.out(format, args);
        }
    }

    private static void processImg(
        Context context,
        Element t,
        Matcher match,
        String startName,
        String format) {
        startName = startName.substring(0, match.start());
        Object value = getValue(context, startName, format);
        String url = null;
        if (value instanceof String) {
            url = (String) value;
            if (url.isEmpty()) {
                value = null;
            }
        }
        if (value == null) {
            U.logError("%s 没有属性值", startName);
            return;
        }
        out("startName=%s; value=%s;", startName, value);
        String relId = null;
        NodeList list = t.getElementsByTagName("v:imagedata");
        Element useElement = null;
        String attrName = null;
        if (list.getLength() > 0) {
            useElement = (Element) list.item(0);
            relId = useElement.getAttribute("r:id");
            attrName = "r:id";
        } else {
            list = t.getElementsByTagName("a:blip");
            if (list.getLength() > 0) {
                useElement = (Element) list.item(0);
                relId = useElement.getAttribute("r:embed");
                attrName = "r:embed";
            }
        }
        if (relId != null) {
            if (context.relationUseMap.get(relId) != null) { // 重复使用的图片
                Element ref = (Element) context.relationMap.get(relId);
                Element e = (Element) ref.cloneNode(true);
                ref.getParentNode().insertBefore(e, ref);
                context.relationMaxId++;
                relId = "rId" + context.relationMaxId;
                e.setAttribute("Id", relId);
                e.setAttribute("Target", ref.getAttribute("Target").replaceAll("\\d+", context.relationMaxId + ""));
                //
                useElement.setAttribute(attrName, relId);
                context.relationMap.put(relId, e);
            }
            context.relationUseMap.put(relId, 1);
            byte[] data = new byte[0];
            if (url != null) {
                if (url.startsWith("http")) {
                    SimpleResponse resp = HttpUtils.doGetBase(url, null);
                    if (resp.data != null && resp.data.length > 0) {
                        data = resp.data;
                    } else {
                        U.logError("下载失败: %s", url);
                    }
                } else {
                    try {
                        data = Files.readAllBytes(new File(url).toPath());
                    } catch (Exception e) {
                        U.logError("读取失败: %s", url);
                    }
                }
            } else if (value instanceof byte[]) {
                data = (byte[]) value;
            }
            if (data != null) {
                Element e = context.relationMap.get(relId);
                String target = e.getAttribute("Target");
                target = "word/" + target;
                context.fileMap.put(target, data);
            }
        } else {
            U.logError("%s 没找到图片引用Id", startName);
        }
    }
}
