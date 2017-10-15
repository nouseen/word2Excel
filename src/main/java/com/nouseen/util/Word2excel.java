package com.nouseen.util;

import com.nouseen.bean.CheckContent;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by nouseen on 2017/9/10.
 */
public class Word2excel {


    private static String testContent;

    public static void extractContenFromExcel() throws IOException {

        // 输出文件
        String outPutPath = "D:\\result.txt";

        // 输出流
        File outPutFile = new File(outPutPath);

        // 写文件流
        FileWriter fileWriter = new FileWriter(outPutFile);

        // 源
        String filePath = "D:\\";

        // 拿到目录下的所有的文件
        File file = new File(filePath);
        File[] tempList = file.listFiles();

        // 遍历所有文件，拿到所有文档
        for (File file1 : tempList) {
            if (file1.getName().contains("docx")) {
                XWPFDocument docx = new XWPFDocument(POIXMLDocument.openPackage(file1.getAbsolutePath()));
                XWPFWordExtractor we = new XWPFWordExtractor(docx);

                docx.getBodyElements();
                // 拿到文档内容
                String text = we.getText();

                // 输出到文件
                fileWriter.append(text);
                fileWriter.flush();
            }
        }


        fileWriter.close();
    }

    /**
     * 通过text的方法获取值
     *
     * @param docx
     * @return
     * @throws NoSuchMethodException
     * @throws InvocationTargetException
     * @throws IllegalAccessException
     */
    public static CheckContent dealXWPFDocumentThoughtText(XWPFDocument docx) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        CheckContent checkContent = new CheckContent();
        List<IBodyElement> bodyElements = docx.getBodyElements();

        XWPFWordExtractor we = new XWPFWordExtractor(docx);
        String text = we.getText();

        return checkContent;
    }

    /**
     * 处理中文文档，2007 docx
     *
     * @param docx
     */
    public static CheckContent dealXWPFDocument(XWPFDocument docx) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        // 返回对象
        CheckContent checkContent = new CheckContent();

        List<IBodyElement> bodyElements = docx.getBodyElements();

        XWPFWordExtractor we = new XWPFWordExtractor(docx);

        // 正则处理表格中无法获取到的属性
        getContenNotInCell(checkContent, we.getText());

        // 拿到表格
        IBodyElement iBodyElement = bodyElements.get(0);
        // 拿到全部内容
        IBody body = iBodyElement.getBody();

        // 获取表格
        List<XWPFTable> tables = body.getTables();

        Class<CheckContent> checkContentClass = CheckContent.class;
        Field[] fields = checkContentClass.getDeclaredFields();

        // 遍历每一个表格i
        int index = 0;
        for (XWPFTable table : tables) {

            // 整个表的文本
            String text = table.getText();

            // 拿到每一行
            table.getNumberOfRows();
            List<XWPFTableRow> rows = table.getRows();

            Iterator<XWPFTableRow> xwpfTableRowIterator = rows.iterator();

            // 遍历行
           while (xwpfTableRowIterator.hasNext()){

               XWPFTableRow xwpfTableRow = xwpfTableRowIterator.next();

               List<XWPFTableCell> tableCells = xwpfTableRow.getTableCells();

                // 拿到每一个单元格，遍历单元格
                Iterator<XWPFTableCell> iterator = tableCells.iterator();
                while (iterator.hasNext()) {
                    XWPFTableCell tableCell = iterator.next();
                    String label = tableCell.getText();

                    // 如果内容含有注解，则取得下一个装入对应的属性
                    for (Field field : fields) {
                        // 如果拿到字段名，则进行注入内容
                        String name = field.getName();
                        ContentFieldName annotation = field.getAnnotation(ContentFieldName.class);
                        if (containList(label, annotation.value())) {
                            if (! iterator.hasNext()) {
                                break;
                            }

                            Method methodGet = checkContentClass.getMethod("get" + name, null);

                            // 如果字段不为空，则说明已经注入了
                            if (StringUtils.isNotBlank((String)methodGet.invoke(checkContent,null))) {
                                break;
                            }
                            // 拿到值
                            String value = iterator.next().getText();

                            // 拿到set 方法
                            Method method = checkContentClass.getMethod("set" + name, String.class);
                            // 反射调用
                            method.invoke(checkContent, value);

                            if (name.equals("肝B超")) {
                                String valueS1 = "";
                                for (int i = 0; i < 3; i++) {
                                    valueS1 = iterator.next().getText();
                                }
                                checkContent.set肝B超提示(valueS1);
                            }

                            break;
                        }
                    }
                }
            }
        }

        return checkContent;
    }


    /**
     * 从文档里面注值
     * @param checkContent
     * @param text
     */
    private static void getContenNotInCell(CheckContent checkContent, String text) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {

        Class<CheckContent> checkContentClass = CheckContent.class;

        String[] needToDealFiles = new String[]{
                "配偶职业及健康状况"
                ,"吸烟史"
                ,"饮酒史"
        };

        Field[] fields = checkContentClass.getDeclaredFields();
        for (Field field : fields) {

            // 只处理指定字段
            String name = field.getName();
            if (! Arrays.asList(needToDealFiles).contains(name)) {
                continue;
            }

            HashMap<String, String> valueMap = matchKeyWords(text, name);

            Set<Map.Entry<String, String>> entries = valueMap.entrySet();
            for (Map.Entry<String, String> entry : entries) {
                // 拿到set方法
                Method methodGet = checkContentClass.getMethod("set" + name, String.class);
                methodGet.invoke(checkContent, entry.getValue());
            }
        }

        //吸烟史：经常吸(10 )支/天,共(3 )年　　,饮酒史：
        // 单独处理家族史
        String s = "家族史\\S*([\\S\\s]*?)其它";
        Pattern pattern = Pattern.compile(s);
        Matcher matcher = pattern.matcher(text);
        if (matcher.find()) {
            checkContent.set家族史(matcher.group(1));
        }
        // 单独处理个人生活史
        String liveHisTory = "吸烟史：(.*),饮酒";
         pattern = Pattern.compile(liveHisTory);
         matcher = pattern.matcher(text);
        if (matcher.find()) {
            checkContent.set吸烟史(matcher.group(1));
        }
        // 单独处理饮酒史
        String drinkHisTory = "饮酒史：(.*)";
         pattern = Pattern.compile(drinkHisTory);
         matcher = pattern.matcher(text);
        if (matcher.find()) {
            checkContent.set饮酒史(matcher.group(1));
        }

        // 单独处理饮酒史
        String bearHisTory = "配偶职业及健康状况：(.*)";
         pattern = Pattern.compile(bearHisTory);
         matcher = pattern.matcher(text);
        if (matcher.find()) {
            checkContent.set配偶职业及健康状况(matcher.group(1));
        }
    }


    /**
     * 匹配关键字
     *
     * @param source
     * @return
     */
    public static HashMap<String, String> matchKeyWords(String source,String target) {
        HashMap<String, String> stringStringHashMap = new HashMap<String, String>();
        String s = "(" + target + ")：?\\s*([^\\s,]+)";
        Pattern pattern = Pattern.compile(s);
        Matcher matcher = pattern.matcher(source);
        if (matcher.find()) {
            stringStringHashMap.put(matcher.group(1),matcher.group(2));
        }
        return stringStringHashMap;
    }

    /**
     * 获取列文字

     * @param xwpfTableRow
     * @return
     */
    private static String getRowTxt(XWPFTableRow xwpfTableRow) {
        StringBuilder stringBuilder = new StringBuilder();

        List<XWPFTableCell> tableCells = xwpfTableRow.getTableCells();
        for (XWPFTableCell tableCell : tableCells) {
            stringBuilder.append(tableCell.getText());
            stringBuilder.append(" ");
        }

        return stringBuilder.toString();
    }

    /**
     * 列表中是否包含
     *
     * @param text
     * @param strings
     * @return
     */
    public static boolean containList(String text, String... strings) {

        if (StringUtils.isBlank(text)) {
            return false;
        }

        if (StringUtils.contains(text, "小结")) {
            return false;
        }
        text = text.replaceAll("\\s", "");
        for (String string : strings) {
            if (StringUtils.contains(text, string)) {
                return true;
            }
        }

        return false;
    }


    public static String getTestContent() {
        return testContent;
    }
}


