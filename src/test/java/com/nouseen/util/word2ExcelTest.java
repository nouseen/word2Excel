package com.nouseen.util;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.POIXMLTextExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.junit.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocument1;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.w3c.dom.NodeList;

import javax.print.Doc;
import java.io.*;
import java.util.List;


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

}