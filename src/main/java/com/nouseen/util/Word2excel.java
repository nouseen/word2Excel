package com.nouseen.util;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
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
    public static void dealXWPFDocument(XWPFDocument docx) {
        List<IBodyElement> bodyElements = docx.getBodyElements();
        System.out.println(bodyElements.size());

        // 拿到表格
        IBodyElement iBodyElement = bodyElements.get(0);
        // 拿到全部内容
        IBody body = iBodyElement.getBody();

        // 获取表格
        List<XWPFTable> tables = body.getTables();


     // String[]keyWordList=new String[]{
     //        "编号","姓名","类别","工作单位","单位电话","体检单位","检查日期","单位地址","出生地","居民身份证号码"
     //        ,"个人联系电话","文化程度","职业照射种类","放射线种类","接触放射线工龄","吸烟史","饮酒史","家族史",
     //        "一般状况","	脉率	","收缩压	","舒张压","	身高","体重	","左眼裸视力	","右眼裸视力","	左眼矫正视力	","右眼矫正视力","	色觉",
     //        "右眼眼底","左眼眼底","	肝B超","肝B超提示","心电图","白细胞计数(WBC)","红细胞计数(RBC)","血红蛋白量(HGB)","红细胞压积(HCT)"
     //         ,"平均红细胞体积(MCV)","平均红细胞血红蛋白量(MCH)", "平均红细胞血红蛋白浓度(MCHC)","血小板计数（PLT）","红细胞分布宽度（RDW-SD）"
     //         ,"红细胞分布宽度（RDW-CV）","血小板分布宽度(PDW)","平均血小板体积(MPV)","大型血小板比率(P-LCR)","血小板压积(PCT)","中性粒细胞百分率(NEUT%)"
     //         ,"淋巴细胞百分率(LYMPH%)","单核细胞百分率(MONO%)","嗜酸性粒细胞百分率(EO%)","嗜碱性粒细胞百分率(BASO%)","中性粒细胞数(NEUT#)"
     //         ,"淋巴细胞数(LYMPH#)","单核细胞数(MONO#)","嗜酸性粒细胞数(EO#)","嗜碱性粒细胞数(BASO#)","血糖","丙氨酸氨基转移酶（ALT）","总胆红素（TBIL）"
     //         ,"总蛋白（TP）","白蛋白（ALB)","球蛋白（GLB)", "白/球比值（A/G）","尿素氮（BUN）","肌酐（CREA）","谷酰转肽酶(GGT)","尿白细胞（WBC）"
     //         ,"酮体（KET）","亚硝酸（NIT)", "尿胆原（URO)","胆红素（BIL)","尿蛋白质（PRO）	葡萄糖	尿比重（SG）","隐血（BLD）","酸碱值（PH）"
     //         ,"维C（Vc）", "游离三碘甲状腺原氨酸(FT3)","游离甲状腺素(FT4)","超敏促甲状腺素(TSH)","AFP(化学发光法)","EB病毒壳抗原lgA抗体"
     //         ,"MN-1000微核分析细胞数", "MN-C微核细胞率","MN微核率","LMY淋巴细胞转化率","染色体分析细胞数","畸变细胞率(染色体型畸变）","染色体型畸变率（%）"
     //         ,"双着丝粒染色体率(dic)","环状染色体率(r)", "无着丝粒片段率(ace)","相互易位率(t)","倒位率(inv)","染色单体型畸变率（%）"
     //    };

        // 遍历每一个表格
        for (XWPFTable table : tables) {

            // 整个表的文本
            String text = table.getText();



            // 拿到每一行
            table.getNumberOfRows();
            List<XWPFTableRow> rows = table.getRows();
            // 遍历行
            for (XWPFTableRow xwpfTableRow : rows) {
                // String rowTxt = getRowTxt(xwpfTableRow);
                // if (rowTxt.contains("检查结果") && ! rowTxt.contains("项目名称")) {
                //     break;
                // }

                // if (containList(rowTxt,keyWordList)) {
                //     System.out.println(rowTxt);
                //     System.out.println("\n");
                // }

                List<XWPFTableCell> tableCells = xwpfTableRow.getTableCells();
                // 拿到每一个单元格，遍历单元格
                Iterator<XWPFTableCell> iterator = tableCells.iterator();
                while (iterator.hasNext()) {
                    XWPFTableCell tableCell = iterator.next();
                    tableCell.getText();
                    // 如果内容含有关键字，则取得下一个装入对应的属性

                }
                // for (XWPFTableCell tableCell : tableCells) {
                //     String text1 = tableCell.getText();
                    // if (StringUtils.equals(text1, "编号")) {
                        // System.out.println();
                    // }
                    // System.out.println(text1);
                    // 如果单元格内容为字段名，则获取下一个单元格的内容
                // }

            }
        }
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


