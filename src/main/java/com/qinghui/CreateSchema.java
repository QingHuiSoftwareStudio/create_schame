package com.qinghui;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

/**
 * 2018年11月25日  16时11分
 *
 * @Author 2710393@qq.com
 * 单枪匹马你别怕，一腔孤勇又如何。
 * 这一路，你可以哭，但是你不能怂。
 */
public class CreateSchema {

    private static Sheet sheet = null;
    private static final String EXCEL_PATH = "G:\\develope\\schema.xlsx";
    private static final String FILE_TYPE_1 = "xls";
    private static final String FILE_TYPE_2 = "xlsx";

    public static void main(String[] args) throws Exception {

        File excel = new File(EXCEL_PATH);
        // 判断路径存在且对应的是一个文件
        if(excel.exists() && excel.isFile()) {

            // 获取文件的后缀名
            String[] split = excel.getName().split("\\.");
            Workbook wb = null;
            if(FILE_TYPE_1.equals(split[1])) {
                FileInputStream inputStream = new FileInputStream(excel);
                wb = new HSSFWorkbook(inputStream);
            }else if(FILE_TYPE_2.equals(split[1])) {
                wb = new XSSFWorkbook(excel);
            }else {
                throw new RuntimeException("只支持xls和xlsx格式的文件!");
            }

            // 解析并创建文件
            sheet = wb.getSheetAt(0);
            createFile(wb.getSheetAt(0));
        }
    }

    /**
     * 创建schema.xml文件
     * @param sheet 文件配置信息excel(数据源)
     * @throws IOException
     */
    private static void createFile(Sheet sheet) throws IOException {

        BufferedWriter bw = new BufferedWriter(new FileWriter(new File("G:\\develope\\schema.xml")));
        bw.write("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>");
        bw.newLine();
        bw.write("<schema name=\"example\" version=\"1.5\">");
        bw.newLine();
        bw.write("\t<field name=\"_version_\" type=\"long\" indexed=\"true\" stored=\"true\"/>");
        bw.newLine();

        // 生成field字段
        createField(sheet, bw);
        bw.newLine();

        // 生成dynamicField字段
        createDynamicField(sheet, bw);
        bw.newLine();

        // 生成uniqueKey字段
        createUniqueKey(sheet, bw);
        bw.newLine();

        // 生成copyField字段
        createCopyField(sheet, bw);
        bw.newLine();

        // 生成fieldType字段
        // ik
        createIKFieldType(sheet, bw);
        // 生成其它常用类型字段
        createOtherFieldType(bw);
        bw.newLine();

        bw.write("</schema>");
        bw.close();
    }

    /**
     * 生成其它常用类型字段
     * @param bw
     * @throws IOException
     */
    private static void createOtherFieldType(BufferedWriter bw) throws IOException {
        bw.write("\t<fieldType name=\"string\" class=\"solr.StrField\" sortMissingLast=\"true\" />");
        bw.newLine();
        bw.write("\t<fieldType name=\"boolean\" class=\"solr.BoolField\" sortMissingLast=\"true\"/>");
        bw.newLine();
        bw.write("\t<fieldType name=\"int\" class=\"solr.TrieIntField\" precisionStep=\"0\" positionIncrementGap=\"0\"/>");
        bw.newLine();
        bw.write("\t<fieldType name=\"float\" class=\"solr.TrieFloatField\" precisionStep=\"0\" positionIncrementGap=\"0\"/>");
        bw.newLine();
        bw.write("\t<fieldType name=\"long\" class=\"solr.TrieLongField\" precisionStep=\"0\" positionIncrementGap=\"0\"/>");
        bw.newLine();
        bw.write("\t<fieldType name=\"double\" class=\"solr.TrieDoubleField\" precisionStep=\"0\" positionIncrementGap=\"0\"/>");
        bw.newLine();
        bw.write("\t<fieldType name=\"date\" class=\"solr.TrieDateField\" precisionStep=\"0\" positionIncrementGap=\"0\"/>");
        bw.newLine();
        bw.write("\t<fieldType name=\"text_general\" class=\"solr.TextField\" positionIncrementGap=\"100\">");
        bw.newLine();
        bw.write("\t\t<analyzer type=\"index\">");
        bw.newLine();
        bw.write("\t\t\t<tokenizer class=\"solr.StandardTokenizerFactory\"/>");
        bw.newLine();
        bw.write("\t\t\t<filter class=\"solr.StopFilterFactory\" ignoreCase=\"true\" words=\"stopwords.txt\" />");
        bw.newLine();
        bw.write("\t\t\t<filter class=\"solr.LowerCaseFilterFactory\"/>");
        bw.newLine();
        bw.write("\t\t</analyzer>");
        bw.newLine();
        bw.write("\t\t<analyzer type=\"query\">");
        bw.newLine();
        bw.write("\t\t\t<tokenizer class=\"solr.StandardTokenizerFactory\"/>");
        bw.newLine();
        bw.write("\t\t\t<filter class=\"solr.StopFilterFactory\" ignoreCase=\"true\" words=\"stopwords.txt\" />");
        bw.newLine();
        bw.write("\t\t\t<filter class=\"solr.SynonymFilterFactory\" synonyms=\"synonyms.txt\" ignoreCase=\"true\" expand=\"true\"/>");
        bw.newLine();
        bw.write("\t\t\t<filter class=\"solr.LowerCaseFilterFactory\"/>");
        bw.newLine();
        bw.write("\t\t</analyzer>");
        bw.newLine();
        bw.write("\t</fieldType>");
    }

    /**
     * 生成其ik字段
     * @param sheet 文件配置信息excel(数据源)
     * @param bw
     * @throws IOException
     */
    private static void createIKFieldType(Sheet sheet, BufferedWriter bw) throws IOException {
        Row rowFTField = sheet.getRow(7);
        if(rowFTField != null) {
            Cell cellFTField = rowFTField.getCell(23);
            if(!checkCellIsNull(cellFTField)) {
                StringBuffer sb = new StringBuffer();
                sb.append("\t<fieldType ");
                sb = appendStr(sb, "name", cellFTField, 7, 23);
                sb.append("class=\"solr.TextField\">");
                bw.write(sb.toString());
                bw.newLine();
                bw.write("\t\t<analyzer type=\"index\" class=\"org.wltea.analyzer.lucene.IKAnalyzer\" />");
                bw.newLine();
                bw.write("\t\t<analyzer type=\"query\" class=\"org.wltea.analyzer.lucene.IKAnalyzer\" />");
                bw.newLine();
                bw.write("\t</fieldType>");
            }
            bw.newLine();
        }
    }

    /**
     * 生成copyField字段
     * @param sheet 文件配置信息excel(数据源)
     * @param bw
     * @throws IOException
     */
    private static void createCopyField(Sheet sheet, BufferedWriter bw) throws IOException {
        for (int start = 7;start < sheet.getLastRowNum(); start++) {
            Row rowCFSource = sheet.getRow(start);
            if(rowCFSource == null) {
                break;
            }
            Cell cellCFSource = rowCFSource.getCell(19);
            if(checkCellIsNull(cellCFSource)) {
                break;
            }

            StringBuffer sb = new StringBuffer();
            sb.append("\t<copyField ");
            sb = appendStr(sb, "source", cellCFSource, start, 19);

            Cell cellCFDest = sheet.getRow(start).getCell(20);
            sb = appendStr(sb, "dest", cellCFDest, start, 20);

            sb.append("/>");
            bw.write(sb.toString());
            bw.newLine();
        }
    }

    /**
     * 生成uniqueKey字段
     * @param sheet 文件配置信息excel(数据源)
     * @param bw
     * @throws IOException
     */
    private static void createUniqueKey(Sheet sheet, BufferedWriter bw) throws IOException {
        Row rowUKField = sheet.getRow(7);
        if(rowUKField != null) {
            Cell cellUKField = rowUKField.getCell(16);
            if(checkCellIsNull(cellUKField)) {
                bw.write("\t<uniqueKey>id</uniqueKey>");
            }else {
                bw.write("\t<uniqueKey>" + cellUKField.getStringCellValue() + "</uniqueKey>");
            }
            bw.newLine();
        }
    }

    /**
     * 生成dynamicField字段
     * @param sheet 文件配置信息excel(数据源)
     * @param bw
     * @throws IOException
     */
    private static void createDynamicField(Sheet sheet, BufferedWriter bw) throws IOException {
        for (int start = 7;start < sheet.getLastRowNum(); start++) {
            Row rowDynamicField = sheet.getRow(start);
            if(rowDynamicField == null) {
                break;
            }
            Cell cellDynamicField = rowDynamicField.getCell(9);
            if(checkCellIsNull(cellDynamicField)) {
                break;
            }

            StringBuffer sb = new StringBuffer();
            sb.append("\t<dynamicField ");
            sb = appendStr(sb, "name", cellDynamicField, start, 9);

            Cell cellDYType = sheet.getRow(start).getCell(10);
            sb = appendStr(sb, "type", cellDYType, start, 10);

            Cell cellDYIndexed = sheet.getRow(start).getCell(11);
            sb = appendStr(sb, "indexed", cellDYIndexed, start, 11);

            Cell cellDYStored = sheet.getRow(start).getCell(12);
            sb = appendStr(sb, "stored", cellDYStored, start, 12);

            Cell cellDYMultiValued = sheet.getRow(start).getCell(13);
            sb = appendStr(sb, "multiValued", cellDYMultiValued, start, 13);

            sb.append("/>");
            bw.write(sb.toString());
            bw.newLine();
        }
    }

    /**
     * 生成field字段
     * @param sheet 文件配置信息excel(数据源)
     * @param bw
     * @throws IOException
     */
    private static void createField(Sheet sheet, BufferedWriter bw) throws IOException {
        for (int start = 7;start < sheet.getLastRowNum(); start++) {
            Row rowField = sheet.getRow(start);
            if(rowField == null) {
                break;
            }
            Cell cellField = rowField.getCell(1);
            if(checkCellIsNull(cellField)) {
                break;
            }

            StringBuffer sb = new StringBuffer();
            sb.append("\t<field ");
            sb = appendStr(sb, "name", cellField, start, 1);

            Cell cellFieldType = sheet.getRow(start).getCell(2);
            sb = appendStr(sb, "type", cellFieldType, start, 2);

            Cell cellFieldIndexed = sheet.getRow(start).getCell(3);
            sb = appendStr(sb, "indexed", cellFieldIndexed, start, 3);

            Cell cellFieldStored = sheet.getRow(start).getCell(4);
            sb = appendStr(sb, "stored", cellFieldStored, start, 4);

            Cell cellFieldRequired = sheet.getRow(start).getCell(5);
            sb = appendStr(sb, "required", cellFieldRequired, start, 5);

            Cell cellMultiValued = sheet.getRow(start).getCell(6);
            sb = appendStr(sb, "multiValued", cellMultiValued, start, 6);

            sb.append("/>");
            bw.write(sb.toString());
            bw.newLine();
        }
    }

    /**
     * 拼接属性字符串
     * @param sb
     * @param attrName 属性名称
     * @param cell 属性值对应的单元格对象
     * @param rowIndex 行索引
     * @param cellIndex 列索引
     * @return
     */
    private static StringBuffer appendStr(StringBuffer sb, String attrName, Cell cell, int rowIndex, int cellIndex) {
        if(checkCellIsNull(cell)) {
            sb.append(attrName).append("=\"").append(sheet.getRow(6).getCell(cellIndex).getStringCellValue()).append("\" ");
        }else {
            sb.append(attrName).append("=\"").append(cell.getStringCellValue()).append("\" ");
        }
        return sb;
    }

    /**
     * 判断单元格对象是否为空
     * @param cell
     * @return
     * true 空
     * false 非空
     */
    private static boolean checkCellIsNull(Cell cell) {
        return cell == null || "".equals(cell.getStringCellValue().trim());
    }

}
