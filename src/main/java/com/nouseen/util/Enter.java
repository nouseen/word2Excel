package com.nouseen.util;

import com.nouseen.bean.CheckContent;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
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
public class Enter {

    public static void main(String[] args) throws IOException, NoSuchMethodException, IllegalAccessException, InvocationTargetException {

        // 待输出列表
        List<CheckContent> checkContentList = new ArrayList<CheckContent>();
        // sheet名字
        String sheetName = "resultExcel";

        // 源
        String filePath = "D:\\";

        // 拿到目录下的所有的文件
        File file=new File(filePath);

        // 文件列表
        File[] tempList = file.listFiles();

        // 遍历所有文件，拿到所有文档
        for (File file1 : tempList) {
            if (file1.getName().contains("docx")) {
                // 打开文件
                XWPFDocument docx = new XWPFDocument(POIXMLDocument.openPackage(file1.getAbsolutePath()));
                // 解析为对像
                CheckContent checkContent = Word2excel.dealXWPFDocument(docx);
                // 添加到待处理列表
                checkContentList.add(checkContent);
            }
        }

        // 拿到对像字段名，用于做Excel的列名
        Field[] declaredFields = CheckContent.class.getDeclaredFields();

        // 装入标题map
        Map<String,String> titleMap = new LinkedHashMap<String,String>();

        for (Field declaredField : declaredFields) {
            // excel 列名map
            titleMap.put(declaredField.getName(), declaredField.getName());
        }

        // 导出Excel
        ExcelUtil.excelExport(checkContentList, titleMap, sheetName);

    }
}
