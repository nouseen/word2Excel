package com.nouseen.util;

import com.nouseen.bean.CheckContent;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;


/**
 * Created by nouseen on 2017/9/10.
 */
public class word2ExcelTest {

    @Test
    public void test() {
        System.out.println("tesst");
    }


    @Test
    public void test201() throws IOException {
        // 输出流
        File outPutFile = new File("D:\\result.text");
        FileWriter fileWriter = new FileWriter(outPutFile);
        // 源
        String filePath = "D:\\";

        // 拿到目录下的所有的文件
        File file=new File(filePath);
        File[] tempList = file.listFiles();
        // 遍历所有文件，拿到所有文档
        for (File file1 : tempList) {
            if (file1.getName().contains("docx")) {
                XWPFDocument docx = new XWPFDocument(POIXMLDocument.openPackage(file1.getAbsolutePath()));
                XWPFWordExtractor we = new XWPFWordExtractor(docx);

                // 拿到文档内容
                String text = we.getText();

                // 输出到文件
                fileWriter.append(text);
                fileWriter.flush();
            }
        }


        fileWriter.close();
    }
    
    @Test
    public void testDeal() throws IOException, NoSuchMethodException, IllegalAccessException, InvocationTargetException {

        String sheetName = "resultExcel";

        XWPFDocument docx = new XWPFDocument(POIXMLDocument.openPackage("D:\\白书昌.xml"));
        CheckContent checkContent = Word2excel.dealXWPFDocument(docx);


        List<CheckContent> checkContentList = new ArrayList<CheckContent>();
        checkContentList.add(checkContent);

        Field[] declaredFields = CheckContent.class.getDeclaredFields();

        // 装入标题map
        Map<String,String> titleMap = new LinkedHashMap<String,String>();
        for (Field declaredField : declaredFields) {
            titleMap.put(declaredField.getName(), declaredField.getName());
        }

        long start = System.currentTimeMillis();
        ExcelUtil.excelExport(checkContentList, titleMap, sheetName);
        long end = System.currentTimeMillis();
        System.out.println("end导出");
        System.out.println("耗时："+(end-start)+"ms");

        System.out.println(checkContent);

    }

    @Test
    public void testDealXml() throws IOException, NoSuchMethodException, IllegalAccessException, InvocationTargetException {
        File file = new File("D:\\白书昌.xml");
        FileInputStream fileInputStream = new FileInputStream(file);

        XWPFDocument docx = new XWPFDocument(fileInputStream);
        Word2excel.dealXWPFDocument(docx);
    }

    @Test
    public void testContain(){
        boolean 姓名 = Word2excel.containList("姓  名", "姓名");
        System.out.println(姓名);

    }
}