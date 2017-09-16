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
     * 处理中文文档
     *
     * @param docx
     */
    public static CheckContent dealXWPFDocument(XWPFDocument docx) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        // 返回对象
        CheckContent checkContent = new CheckContent();

        List<IBodyElement> bodyElements = docx.getBodyElements();

        XWPFWordExtractor we = new XWPFWordExtractor(docx);

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

            // System.out.println(String.format("当前第%s段\n", ++index));
            // System.out.println(text);

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


                    // 如果内容含有关键字，则取得下一个装入对应的属性
                    for (Field field : fields) {
                        // 如果拿到字段名，则进行注入内容
                        String name = field.getName();

                        if (containList(label, name)) {
                            if (! iterator.hasNext()) {
                                break;
                            }

                            // 拿到get方法
                            if (label.contains("总")) {
                                if (name.contains("总")) {

                                } else {
                                    name = "总" + name;
                                }
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
        return "\"C:\\Program Files\\Java\\jdk1.8.0_73\\bin\\java\" -ea -Didea.test.cyclic.buffer.size=1048576 \"-javaagent:C:\\Program Files\\JetBrains\\IntelliJ IDEA 2017.1.1\\lib\\idea_rt.jar=50599:C:\\Program Files\\JetBrains\\IntelliJ IDEA 2017.1.1\\bin\" -Dfile.encoding=UTF-8 -classpath \"C:\\Program Files\\JetBrains\\IntelliJ IDEA 2017.1.1\\lib\\idea_rt.jar;C:\\Program Files\\JetBrains\\IntelliJ IDEA 2017.1.1\\plugins\\junit\\lib\\junit-rt.jar;C:\\Program Files\\Java\\jdk1.8.0_73\\jre\\lib\\charsets.jar;C:\\Program Files\\Java\\jdk1.8.0_73\\jre\\lib\\deploy.jar;C:\\Program Files\\Java\\jdk1.8.0_73\\jre\\lib\\ext\\access-bridge-64.jar;C:\\Program Files\\Java\\jdk1.8.0_73\\jre\\lib\\ext\\cldrdata.jar;C:\\Program Files\\Java\\jdk1.8.0_73\\jre\\lib\\ext\\dnsns.jar;C:\\Program Files\\Java\\jdk1.8.0_73\\jre\\lib\\ext\\jaccess.jar;C:\\Program Files\\Java\\jdk1.8.0_73\\jre\\lib\\ext\\jfxrt.jar;C:\\Program Files\\Java\\jdk1.8.0_73\\jre\\lib\\ext\\localedata.jar;C:\\Program Files\\Java\\jdk1.8.0_73\\jre\\lib\\ext\\nashorn.jar;C:\\Program Files\\Java\\jdk1.8.0_73\\jre\\lib\\ext\\sunec.jar;C:\\Program Files\\Java\\jdk1.8.0_73\\jre\\lib\\ext\\sunjce_provider.jar;C:\\Program Files\\Java\\jdk1.8.0_73\\jre\\lib\\ext\\sunmscapi.jar;C:\\Program Files\\Java\\jdk1.8.0_73\\jre\\lib\\ext\\sunpkcs11.jar;C:\\Program Files\\Java\\jdk1.8.0_73\\jre\\lib\\ext\\zipfs.jar;C:\\Program Files\\Java\\jdk1.8.0_73\\jre\\lib\\javaws.jar;C:\\Program Files\\Java\\jdk1.8.0_73\\jre\\lib\\jce.jar;C:\\Program Files\\Java\\jdk1.8.0_73\\jre\\lib\\jfr.jar;C:\\Program Files\\Java\\jdk1.8.0_73\\jre\\lib\\jfxswt.jar;C:\\Program Files\\Java\\jdk1.8.0_73\\jre\\lib\\jsse.jar;C:\\Program Files\\Java\\jdk1.8.0_73\\jre\\lib\\management-agent.jar;C:\\Program Files\\Java\\jdk1.8.0_73\\jre\\lib\\plugin.jar;C:\\Program Files\\Java\\jdk1.8.0_73\\jre\\lib\\resources.jar;C:\\Program Files\\Java\\jdk1.8.0_73\\jre\\lib\\rt.jar;C:\\development\\JavaProject\\word2Excel\\target\\test-classes;C:\\development\\JavaProject\\word2Excel\\target\\classes;C:\\Users\\nouseen\\.m2\\repository\\junit\\junit\\4.12\\junit-4.12.jar;C:\\Users\\nouseen\\.m2\\repository\\org\\hamcrest\\hamcrest-core\\1.3\\hamcrest-core-1.3.jar;C:\\Users\\nouseen\\.m2\\repository\\org\\apache\\poi\\poi-ooxml\\3.15\\poi-ooxml-3.15.jar;C:\\Users\\nouseen\\.m2\\repository\\org\\apache\\poi\\poi-ooxml-schemas\\3.15\\poi-ooxml-schemas-3.15.jar;C:\\Users\\nouseen\\.m2\\repository\\org\\apache\\xmlbeans\\xmlbeans\\2.6.0\\xmlbeans-2.6.0.jar;C:\\Users\\nouseen\\.m2\\repository\\stax\\stax-api\\1.0.1\\stax-api-1.0.1.jar;C:\\Users\\nouseen\\.m2\\repository\\com\\github\\virtuald\\curvesapi\\1.04\\curvesapi-1.04.jar;C:\\Users\\nouseen\\.m2\\repository\\org\\apache\\poi\\poi\\3.16\\poi-3.16.jar;C:\\Users\\nouseen\\.m2\\repository\\commons-codec\\commons-codec\\1.10\\commons-codec-1.10.jar;C:\\Users\\nouseen\\.m2\\repository\\org\\apache\\commons\\commons-collections4\\4.1\\commons-collections4-4.1.jar;C:\\Users\\nouseen\\.m2\\repository\\org\\apache\\poi\\poi-scratchpad\\3.0.2-FINAL\\poi-scratchpad-3.0.2-FINAL.jar;C:\\Users\\nouseen\\.m2\\repository\\commons-logging\\commons-logging\\1.1\\commons-logging-1.1.jar;C:\\Users\\nouseen\\.m2\\repository\\log4j\\log4j\\1.2.13\\log4j-1.2.13.jar;C:\\Users\\nouseen\\.m2\\repository\\org\\apache\\commons\\commons-lang3\\3.0\\commons-lang3-3.0.jar\" com.intellij.rt.execution.junit.JUnitStarter -ideVersion5 com.nouseen.util.word2ExcelTest,testDeal\n" +
                "\n" +
                "编号：\t201510003616\n" +
                "类别：\t在岗期间\n" +
                "\n" +
                "\n" +
                "\n" +
                "放射工作人员职业健康检查表\n" +
                "\n" +
                "\n" +
                "\n" +
                "\n" +
                "姓    名：\t白书昌\n" +
                "工作单位：\t佛山市第二人民医院\n" +
                "单位电话：\t88032266\n" +
                "体检单位：\t佛山市职业病防治所\n" +
                "检查日期：\t2015-08-18\n" +
                "\n" +
                "\n" +
                "\n" +
                "\n" +
                "登记编号:201510003616 　　　　　　　　　白书昌 　　　男  　　30岁         第 1 页 共 9 页\n" +
                "单位名称:佛山市职业病防治所                                 地址:佛山市禅城区影荫路3号公卫大楼综合楼2楼\n" +
                "资质证书编号: 粤职健协职检2014018号                            电话:83374614\n" +
                "\n" +
                "登记\n" +
                "单位地址：\t佛山市禅城区卫国路78号\n" +
                "邮政编码：\t\t联系人：\t\t电话：\t88032266\n" +
                "\n" +
                "\n" +
                "（个人基本资料）\n" +
                "姓 名：\t白书昌\t性别：\t男\t年    龄：\t30\n" +
                "出生地：\t河南\t民族：\t汉族\t职务/职称：\t主治医师\n" +
                "居民身份证号码：\t411330198507234817\n" +
                "工作单柆：\t佛山市第二人民医院\n" +
                "家庭地址：\t广东省佛山市禅城区卫国路78号\t邮政编码：\n" +
                "个人联系电话：\t18927701054\t文化程度：\t研究生以上\n" +
                "职业照射种类：\t介入放射学2E\t放射线种类：\tγ\n" +
                "接触放射线工龄：\t4年\t每日接触放射线时间：\t8小时\n" +
                "\n" +
                "\n" +
                "\n" +
                "婚姻史\n" +
                "结婚日期：    配偶接触放射线情况：  无\n" +
                "\n" +
                "配偶职业及健康状况：   良好\n" +
                "\n" +
                "生育史\n" +
                "现有子女： 女孩（ 1）人人, ， 子女健康情况： 良好 ， 流产：  次，畸胎：  次，多胎：  次，异位妊娠：  次，不孕不育原因：  -\n" +
                "\n" +
                "个人生活史（长期生活地区，饮食习惯，有无地方病流行地区或疫区生活史、药物滥用情况及烟酒嗜好等）\n" +
                "   吸烟史：无　　,饮酒史：偶饮酒\n" +
                "\n" +
                "家族史（家族中有无遗传性疾病、血液病、糖尿病、高血压病，神经精神性疾病，肿瘤，结核病等）\n" +
                "    无\n" +
                "\n" +
                "其它\n" +
                "\n" +
                " 症状: +表示有症状,-表示无症状\n" +
                "症状\t             负责医生：\n" +
                "\n" +
                "项目名称\t结果\t项目名称\t结果\t项目名称\t结果\t项目名称\t结果\n" +
                "头痛\t－\t羞明\t－\t气短\t－\t便秘\t－\n" +
                "头（晕）昏\t－\t流泪\t－\t胸闷\t－\t尿频\t－\n" +
                "眩晕\t－\t嗅觉减退\t－\t胸痛\t－\t尿急\t－\n" +
                "失眠\t－\t鼻干燥\t－\t咳嗽\t－\t尿血\t－\n" +
                "嗜睡\t－\t鼻塞\t－\t咳痰\t－\t皮下出血\t－\n" +
                "多梦\t－\t流鼻血\t－\t咯血\t－\t皮肤瘙痒\t－\n" +
                "记忆力减退\t－\t流涕\t－\t哮喘\t－\t皮疹\t－\n" +
                "易激动\t－\t耳鸣\t－\t心悸\t－\t浮肿\t－\n" +
                "疲乏无力\t－\t耳聋\t－\t心前区不适\t－\t脱发\t－\n" +
                "低热\t－\t口渴\t－\t食欲减退\t－\t关节痛\t－\n" +
                "盗汗\t－\t流涎\t－\t消瘦\t－\t肌肉酸痛\t－\n" +
                "多汗\t－\t牙痛\t－\t恶心\t－\t肌肉抽搐\t－\n" +
                "全身酸痛\t－\t牙齿松动\t－\t呕吐\t－\t四肢麻木\t－\n" +
                "性欲减退\t－\t刷牙出血\t－\t腹胀\t－\t动作不灵活\t－\n" +
                "视物模糊\t－\t口腔异味\t－\t腹痛\t－\n" +
                "视力下降\t－\t口腔溃疡\t－\t肝区痛\t－\n" +
                "眼痛\t－\t咽痛\t－\t腹泻\t－\n" +
                "\n" +
                "体检编号                                                           第 1 页 共 10 页\n" +
                "单位名称:佛山市职业病防治所                                 地址:佛山市禅城区影荫路3号公卫大楼综合楼2楼\n" +
                "资质证书编号: 粤卫职检2004024号                            电话:83374614\n" +
                "\n" +
                "一般情况\t负责医生:\n" +
                "\n" +
                "项目名称\t检查结果\t单位\t参考范围\t提示\n" +
                "一般状况\t良好\n" +
                "脉率\t100\t次/分\t60-100\n" +
                "收缩压\t136\tmmHg\t90-139\n" +
                "舒张压\t85\tmmHg\t60-89\n" +
                "身高\t169\tcm\n" +
                "体重\t77\tkg\n" +
                "          小结：\t未见异常\n" +
                "视力\t负责医生:\n" +
                "\n" +
                "项目名称\t检查结果\t单位\t参考范围\t提示\n" +
                "左眼裸视力\t-\n" +
                "右眼裸视力\t-\n" +
                "左眼矫正视力\t5.1\n" +
                "右眼矫正视力\t5.1\n" +
                "色觉\t-\n" +
                "          小结：\t左眼矫正视力5.1,右眼矫正视力5.1\n" +
                "内科\t负责医生:\n" +
                "\n" +
                "项目名称\t检查结果\t单位\t参考范围\t提示\n" +
                "发育\t正常\n" +
                "营养\t正常\n" +
                "心脏\t正常\n" +
                "肺\t正常\n" +
                "腹部\t正常\n" +
                "肝\t正常\n" +
                "胆\t正常\n" +
                "脾\t正常\n" +
                "肾\t正常\n" +
                "浅表淋巴结\t正常\n" +
                "甲状腺\t正常\n" +
                "脊柱\t正常\n" +
                "四肢\t正常\n" +
                "其它\t-\n" +
                "          小结：\t未见异常\n" +
                "皮肤\t负责医生:\n" +
                "\n" +
                "项目名称\t检查结果\t单位\t参考范围\t提示\n" +
                "皲裂\t未发现\n" +
                "出血紫癜\t未发现\n" +
                "干燥\t未发现\n" +
                "色素减退\t未发现\n" +
                "过度角化\t未发现\n" +
                "皮肤萎缩\t未发现\n" +
                "脱毛,脱发\t未发现\n" +
                "皮疹\t未发现\n" +
                "色素沉着\t未发现\n" +
                "多汗\t未发现\n" +
                "溃疡\t未发现\n" +
                "疣状物\t未发现\n" +
                "脱屑\t未发现\n" +
                "指甲\t正常\n" +
                "          小结：\t未见异常\n" +
                "神经系统\t负责医生:\n" +
                "\n" +
                "项目名称\t检查结果\t单位\t参考范围\t提示\n" +
                "皮肤划痕症\t-\n" +
                "膝反射\t存在\n" +
                "跟腱反射\t存在\n" +
                "肌力\t正常\n" +
                "肌张力\t正常\n" +
                "共济运动\t协调\n" +
                "感觉异常\t无\n" +
                "三颤\t无\n" +
                "病理反射\t未引出\n" +
                "          小结：\t未见异常\n" +
                "眼科\t负责医生:\n" +
                "\n" +
                "项目名称\t检查结果\t单位\t参考范围\t提示\n" +
                "右眼晶体\t正常\n" +
                "左眼晶体\t正常\n" +
                "右眼玻璃体\t正常\n" +
                "左眼玻璃体\t正常\n" +
                "右眼眼底\t视乳头边界清,网膜平伏,C/D≈0.3\n" +
                "左眼眼底\t视乳头边界清,网膜平伏,C/D≈0.3\n" +
                "          小结：\t未见异常\n" +
                "B超\t负责医生:\n" +
                "\n" +
                "项目名称\t检查结果\t单位\t参考范围\t提示\n" +
                "肝B超\t脂肪肝（轻 ）\t\t\t肝右叶斜径(15.6 )cm，回声不均匀，细密；边清，前段回声增强，后段声衰减。\n" +
                "胆B超\t 未见异常 \t\t\t大小正常，壁光滑，未见异常回声。\n" +
                "脾B超\t 未见异常 \t\t\t大小正常，实质回声均匀。\n" +
                "双肾B超\t 未见异常 \t\t\t大小形态正常，未见异常回声。\n" +
                "膀胱B超\t 未见异常 \t\t\t膀胱充盈，壁光滑，未见明显异常回声。\n" +
                "前列腺B超\t 未见异常 \t\t\t大小正常，回声均匀，未见异常回声团。\n" +
                "          小结：\t脂肪肝（轻 ）\n" +
                "心电图\t负责医生:\n" +
                "\n" +
                "项目名称\t检查结果\t单位\t参考范围\t提示\n" +
                "心电图\t正常心电图\n" +
                "          小结：\t未见异常\n" +
                "X线检查\t负责医生:\n" +
                "\n" +
                "项目名称\t检查结果\t单位\t参考范围\t提示\n" +
                "X线胸片\t双肺及心膈未见明显异常\n" +
                "          小结：\t未见异常\n" +
                "血常规\t负责医生:\n" +
                "\n" +
                "项目名称\t检查结果\t单位\t参考范围\t提示\n" +
                "白细胞计数(WBC)\t7.07\tx10^9/L\t4~10\n" +
                "红细胞计数(RBC)\t5.83\tx10^12/L\t3.5-5.5\n" +
                "血红蛋白量(HGB)\t172.00\tg/L\t110-172\n" +
                "红细胞压积(HCT)\t40.50\t\t33.5-50.8\n" +
                "平均红细胞体积(MCV)\t69.50\tfL\t70-99.1\n" +
                "平均红细胞血红蛋白量(MCH)\t29.50\tpg\t26.9-33.8\n" +
                "平均红细胞血红蛋白浓度(MCHC)\t425.00\tg/L\t310-370\n" +
                "血小板计数（PLT）\t207.00\tx10^9/L\t100-300\n" +
                "红细胞分布宽度（RDW-SD）\t32.70\tfL\t39-53.9\n" +
                "红细胞分布宽度（RDW-CV）\t13.00\t\t11.9-14.5\n" +
                "血小板分布宽度(PDW)\t8.60\tfL\t9.8-16.2\n" +
                "平均血小板体积(MPV)\t7.90\tfL\t9.4-12.6\n" +
                "大型血小板比率(P-LCR)\t10.70\t\t19.1-47\n" +
                "血小板压积(PCT)\t0.16\t%\t0.16-0.38\n" +
                "中性粒细胞百分率(NEUT%)\t58.34\t\t50-70\n" +
                "淋巴细胞百分率(LYMPH%)\t34.84\t\t20-40\n" +
                "单核细胞百分率(MONO%)\t4.54\t\t3-10\n" +
                "嗜酸性粒细胞百分率(EO%)\t2.34\t%\t0.5-5\n" +
                "嗜碱性粒细胞百分率(BASO%)\t0.14\t\t0-1\n" +
                "中性粒细胞数(NEUT#)\t4.12\tx10^9/L\t2-7\n" +
                "淋巴细胞数(LYMPH#)\t2.46\tx10^9/L\t0.8-4\n" +
                "单核细胞数(MONO#)\t0.32\tx10^9/L\t0.12-1\n" +
                "嗜酸性粒细胞数(EO#)\t0.16\tx10^9/L\t0.02-0.5\n" +
                "嗜碱性粒细胞数(BASO#)\t0.01\tx10^9/L\t0~0.1\n" +
                "          小结：\t红细胞计数(RBC)5.83x10^12/L偏高,平均红细胞体积(MCV)69.50fL偏低,平均红细胞血红蛋白浓度(MCHC)425.00g/L偏高,红细胞分布宽度（RDW-SD）32.70fL偏低,血小板分布宽度(PDW)8.60fL偏低,平均血小板体积(MPV)7.90fL偏低,大型血小板比率(P-LCR)10.70偏低\n" +
                "生化\t负责医生:\n" +
                "\n" +
                "项目名称\t检查结果\t单位\t参考范围\t提示\n" +
                "血糖\t4.8\tmmol/L\t3.89~6.11\n" +
                "丙氨酸氨基转移酶（ALT）\t39\tIU/L\t5~40\n" +
                "总胆红素（TBIL）\t10.8\tUmol/L\t3.42-20.5\n" +
                "总蛋白（TP）\t71.5\tg/L\t60~83\n" +
                "白蛋白（ALB)\t49.1\tg/L\t35~52\n" +
                "球蛋白（GLB)\t22.4\tg/L\t20~30\n" +
                "白/球比值（A/G）\t2.19\t\t1.1~2.5:1\n" +
                "尿素氮（BUN）\t6.3\tmmol/L\t1.7-8.3\n" +
                "肌酐（CREA）\t97\tumol/L\t59-104\n" +
                "谷酰转肽酶(GGT)\t81\tU/L\t8-58\n" +
                "          小结：\t谷酰转肽酶(GGT)81U/L偏高\n" +
                "尿常规\t负责医生:\n" +
                "\n" +
                "项目名称\t检查结果\t单位\t参考范围\t提示\n" +
                "尿白细胞（WBC）\t-\t\t0~10\n" +
                "酮体（KET）\t+1\t\t阴性\n" +
                "亚硝酸（NIT)\t-\t\t阴性\n" +
                "尿胆原（URO)\tNormal\tumol/l\t0~16\n" +
                "胆红素（BIL)\t-\t\t阴性\n" +
                "尿蛋白质（PRO）\t-\t\t阴性\n" +
                "葡萄糖\t-\t\t阴性\n" +
                "尿比重（SG）\t1.030\t\t1.015~1.025\n" +
                "隐血（BLD）\t-\t\t阴性\n" +
                "酸碱值（PH）\t5.0\t\t4.5~8.0\n" +
                "维C（Vc）\t+-\t\t阴性\n" +
                "          小结：\t酮体（KET）+1,尿比重（SG）1.030偏高,酸碱值（PH）5.0偏低,维C（Vc）+-\n" +
                "实验室检查二\t负责医生:\n" +
                "\n" +
                "项目名称\t检查结果\t单位\t参考范围\t提示\n" +
                "游离三碘甲状腺原氨酸(FT3)\t5.09\tpmol/L\t3.10-6.80\n" +
                "游离甲状腺素(FT4)\t19.83\tpmol/L\t12.00-22.00\n" +
                "超敏促甲状腺素(TSH)\t0.753\tuIU/mL\t0.270-4.200\n" +
                "AFP(化学发光法)\t4.91\tng/ml\t≤8.78\n" +
                "EB病毒壳抗原lgA抗体\t阴性\t\t阴性\n" +
                "MN-1000微核分析细胞数\t1000\t个\n" +
                "MN-C微核细胞率\t0\t‰\t0-6\n" +
                "MN微核率\t0\t‰\t0-6\n" +
                "LMY淋巴细胞转化率\t73\t%\t48-90\n" +
                "          小结：\t未见异常\n" +
                "\n" +
                "\n" +
                "检查结果:\n" +
                "1、视力:左眼矫正视力5.1,右眼矫正视力5.1\t2、B超:脂肪肝（轻 ）\t3、血常规:红细胞计数(RBC)5.83x10^12/L偏高\t4、肝功能:谷酰转肽酶(GGT)81U/L偏高\n" +
                "结论：\n" +
                "本次体检未发现接触电离辐射作业职业禁忌证及疑似职业病。\n" +
                "意见:\n" +
                "1、可继续原放射工作。\t2、脂肪肝、肝功能异常,建议到肝病专科进一步诊治，定期追踪复查，适度锻炼，低脂限酒饮食。\t3、血常规异常，建议复查。\n" +
                "\n" +
                "\n" +
                "总检医生：\t\t总检日期：\t2015-10-15\t盖章：\n" +
                "联系电话：\n" +
                "\n" +
                "\n" +
                "end导出\n" +
                "耗时：432ms\n" +
                "com.nouseen.bean.CheckContent@595b007d\n" +
                "\n" +
                "Process finished with exit code 0\n";
    }

}


