package com.nouseen.util;

import java.io.*;

import org.apache.poi.hpsf.DocumentSummaryInformation;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.hwpf.usermodel.TableRow;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.POIXMLTextExtractor;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

/**
 * Created by nouseen on 2017/9/10.
 */
public class Word2excel {



    public static void mergeCellVertically(XWPFTable table, int col, int fromRow, int toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            CTVMerge vmerge = CTVMerge.Factory.newInstance();
            if (rowIndex == fromRow) {
                // The first merged cell is set with RESTART merge value
                vmerge.setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                vmerge.setVal(STMerge.CONTINUE);
            }
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            // Try getting the TcPr. Not simply setting an new one every time.
            CTTcPr tcPr = cell.getCTTc().getTcPr();
            if (tcPr != null) {
                tcPr.setVMerge(vmerge);
            } else {
                // only set an new TcPr if there is not one already
                tcPr = CTTcPr.Factory.newInstance();
                tcPr.setVMerge(vmerge);
                cell.getCTTc().setTcPr(tcPr);
            }
        }
    }

    public void testWord() {
        try {
            FileInputStream in = new FileInputStream("D:\\sinye.doc");//载入文档
            POIFSFileSystem pfs = new POIFSFileSystem(in);
            HWPFDocument hwpf = new HWPFDocument(pfs);
            org.apache.poi.hwpf.usermodel.Range range = hwpf.getRange();//得到文档的读取范围
            TableIterator it = new TableIterator(range);
            //迭代文档中的表格
            while (it.hasNext()) {
                Table tb = (Table) it.next();
                //迭代行，默认从0开始
                for (int i = 0; i < tb.numRows(); i++) {
                    TableRow tr = tb.getRow(i);
                    //迭代列，默认从0开始
                    for (int j = 0; j < tr.numCells(); j++) {
                        org.apache.poi.hwpf.usermodel.TableCell td = tr.getCell(j);//取得单元格
                        //取得单元格的内容
                        for (int k = 0; k < td.numParagraphs(); k++) {
                            org.apache.poi.hwpf.usermodel.Paragraph para = td.getParagraph(k);
                            String s = para.text();
                            System.out.println(s);
                        } //end for
                    }   //end for
                }   //end for
            } //end while
        } catch (Exception e) {
            e.printStackTrace();
        }
    }//end method


    public static void readDocumentSummary(HWPFDocument doc) {
        DocumentSummaryInformation summaryInfo = doc.getDocumentSummaryInformation();
        String category = summaryInfo.getCategory();
        String company = summaryInfo.getCompany();
        int lineCount = summaryInfo.getLineCount();
        int sectionCount = summaryInfo.getSectionCount();
        int slideCount = summaryInfo.getSlideCount();

        System.out.println("---------------------------");
        System.out.println("Category: " + category);
        System.out.println("Company: " + company);
        System.out.println("Line Count: " + lineCount);
        System.out.println("Section Count: " + sectionCount);
        System.out.println("Slide Count: " + slideCount);

    }

    private void handleUpload(File file) {
        Workbook wb = tryToHandleHSSF(file);
        if (wb == null)
            wb = tryToHandleXSSF(file);
        if (wb != null) {
            // ... do the parsing stuff
        }
    }

    /**
     * helper for HSSF
     */
    public Workbook tryToHandleHSSF(File file) {
        try {
            return new HSSFWorkbook(new FileInputStream(file));
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * helper for XSSF
     */
    public Workbook tryToHandleXSSF(File file) {
        Workbook workbook;
        try {
            InputStream fin = new FileInputStream(file);
            BufferedInputStream in = new BufferedInputStream(fin);
            try {
                if (POIFSFileSystem.hasPOIFSHeader(in)) {
                    // if the file is encrypted
                    POIFSFileSystem fs = new POIFSFileSystem(in);
                    EncryptionInfo info = new EncryptionInfo(fs);
                    Decryptor d = Decryptor.getInstance(info);
                    d.verifyPassword(Decryptor.DEFAULT_PASSWORD);
                    workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook(d.getDataStream(fs));
                } else
                    return new org.apache.poi.xssf.usermodel.XSSFWorkbook(in);
            } finally {
                in.close();
            }
        } catch (Exception e) {
            return null;
        }
        return workbook;
    }

}


