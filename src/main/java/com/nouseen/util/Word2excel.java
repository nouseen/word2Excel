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
import java.util.Iterator;
import java.util.List;

/**
 * Created by nouseen on 2017/9/10.
 */
public class Word2excel {


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
     * 处理中文文档
     *
     * @param docx
     */
    public static CheckContent dealXWPFDocument(XWPFDocument docx) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        List<IBodyElement> bodyElements = docx.getBodyElements();
        System.out.println(bodyElements.size());

        // 拿到表格
        IBodyElement iBodyElement = bodyElements.get(0);
        // 拿到全部内容
        IBody body = iBodyElement.getBody();

        // 获取表格
        List<XWPFTable> tables = body.getTables();

        CheckContent checkContent = new CheckContent();
        Class<CheckContent> checkContentClass = CheckContent.class;
        Field[] fields = checkContentClass.getDeclaredFields();

        // 遍历每一个表格
        for (XWPFTable table : tables) {

            // 整个表的文本
            String text = table.getText();

            // 拿到每一行
            table.getNumberOfRows();
            List<XWPFTableRow> rows = table.getRows();
            // 遍历行
            for (XWPFTableRow xwpfTableRow : rows) {

                List<XWPFTableCell> tableCells = xwpfTableRow.getTableCells();
                // 拿到每一个单元格，遍历单元格
                Iterator<XWPFTableCell> iterator = tableCells.iterator();
                while (iterator.hasNext()) {
                    XWPFTableCell tableCell = iterator.next();
                    String label = tableCell.getText();
                    // 如果内容含有关键字，则取得下一个装入对应的属性
                    for (Field field : fields) {
                        // 如果拿到字段名，则进行注入内容
                        if (containList(label,field.getName())) {
                            if (! iterator.hasNext()) {
                                break;
                            }

                            // 拿到get方法
                            Method methodGet = checkContentClass.getMethod("get" + field.getName(), null);

                            // 如果字段不为空，则说明已经注入了
                            if (StringUtils.isNotBlank((String)methodGet.invoke(checkContent,null))) {
                                System.out.println("已注入");
                                break;
                            }

                            // 拿到值
                            String value = iterator.next().getText();
                            // 拿到set 方法
                            Method method = checkContentClass.getMethod("set" + field.getName(), String.class);
                            // 反射调用
                            method.invoke(checkContent, value);
                        }
                    }
                }
            }
        }

        return checkContent;
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


}


