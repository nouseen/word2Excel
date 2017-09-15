package com.nouseen.util;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;

/**
 * Created by nouseen on 2017/9/10.
 */
public class Enter {

    public static void main(String[] args) throws IOException {
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
