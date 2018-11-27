package com.qinghui;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
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
    private static Workbook  wb = null;
    private static final String EXCEL_PATH = "G:\\develope\\schema.xls";
    private static final String FILE_TYPE_1 = "xls";
    private static final String FILE_TYPE_2 = "xlsx";

    public static void main(String[] args) throws Exception {

        File excel = new File(EXCEL_PATH);
        // 判断路径存在且对应的是一个文件
        if(excel.exists() && excel.isFile()) {

            // 获取文件的后缀名
            String[] split = excel.getName().split("\\.");
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
            createFile();

            // 把提示信息写入文件中
            FileOutputStream out = new FileOutputStream(excel);
            wb.write(out);
            out.close();

        }
    }

    /**
     * 创建schema.xml文件
     * @throws IOException
     */
    private static void createFile() throws IOException {

        BufferedWriter bw = new BufferedWriter(new FileWriter(new File("G:\\develope\\schema.xml")));
        bw.write("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>");
        bw.newLine();
        bw.write("<schema name=\"example\" version=\"1.5\">");
        bw.newLine();
        bw.write("\t<field name=\"_version_\" type=\"long\" indexed=\"true\" stored=\"true\"/>");
        bw.newLine();

        // 生成field字段
        createField(bw);
        bw.newLine();

        // 生成dynamicField字段
        createDynamicField(bw);
        bw.newLine();

        // 生成uniqueKey字段
        createUniqueKey(bw);
        bw.newLine();

        // 生成copyField字段
        createCopyField(bw);
        bw.newLine();

        // 生成自定义fieldType字段
        createIKFieldType2(bw);
        // 生成其它fieldType字段
        createOtherFieldType(bw);

    }

    /**
     * 生成其它常用类型字段
     * @param bw
     * @throws IOException
     */
    private static void createOtherFieldType(BufferedWriter bw) throws IOException {

        try (BufferedReader br = new BufferedReader(new InputStreamReader(CreateSchema.class.getClassLoader().getResourceAsStream("schema.txt"))))
        {
            String readLine = br.readLine();
            while (readLine != null) {
                bw.write(readLine);
                bw.newLine();
                readLine = br.readLine();
            }
        }catch (FileNotFoundException e) {
            throw new RuntimeException("请检查classpath下的schema.txt文件是否存在，且内容未被修改!");
        }
        bw.close();
    }

    /**
     * 生成自定义fieldType字段
     * @param bw
     * @throws IOException
     */
    private static void createIKFieldType2(BufferedWriter bw) throws IOException {
        StringBuffer sb = new StringBuffer();
        for (int start = 7;start < sheet.getLastRowNum(); start++) {
            Row rowFTField = sheet.getRow(start);
            // 判断fieldType字段这一行是否有值
            if(rowFTField != null) {
                // 判断当前fieldType字段是否有未读取的值
                if(checkCellIsNull(rowFTField.getCell(23)) && checkCellIsNull(rowFTField.getCell(26)) && checkCellIsNull(rowFTField.getCell(27)) && checkCellIsNull(rowFTField.getCell(29))) {
                    break;
                }
                // 判断fieldType字段是否输入了name值
                if(checkCellIsNull(rowFTField.getCell(23))) {
                    // 判断analyzer是否有值
                    if(checkCellIsNull(rowFTField.getCell(26))) {
                        // 判断filter是否有值
                        if(checkCellIsNull(rowFTField.getCell(29))) {
                            if(sheet.getRow(rowFTField.getRowNum() + 1) == null) {
                                sb.append("\n\t\t</analyzer>\n\t</fieldType>\r\n\r\n");
                                break;
                            } else if (sheet.getRow(rowFTField.getRowNum() + 1).getCell(29) == null){
                                sb.append("\n\t\t</analyzer>\n\t</fieldType>\r\n\r\n");
                            }
                        }else {
                            // 拼接filter
                            sb.append("\n\t\t\t<filter ");
                            sb = appendFieldTypeStr(sb,"class", rowFTField, 29);
                            sb = appendFieldTypeStr(sb,"ignoreCase", rowFTField, 30);
                            sb = appendFieldTypeStr(sb,"words", rowFTField, 31);
                            sb = appendFieldTypeStr(sb,"format", rowFTField, 32);
                            sb.append("/>");
                            if(sheet.getRow(rowFTField.getRowNum() + 1) == null || sheet.getRow(rowFTField.getRowNum() + 1).getCell(23) != null || sheet.getRow(rowFTField.getRowNum() + 1).getCell(29) == null) {
                                sb.append("\n\t\t</analyzer>\n\t</fieldType>\r\n\r\n");
                            }
                        }
                    }else {
                        // 拼接analyzer
                        sb.append("\n\t\t</analyzer>\n\t\t<analyzer ");
                        sb = appendFieldTypeStr(sb,"type", rowFTField, 26);
                        sb.append(">");
                        // 拼接tokenizer
                        sb.append("\n\t\t\t<tokenizer ");
                        sb = appendFieldTypeStr(sb,"class", rowFTField, 27);
                        sb = appendFieldTypeStr(sb,"mode", rowFTField, 28);
                        sb.append("/>");
                        // 拼接filter
                        sb.append("\n\t\t\t<filter ");
                        sb = appendFieldTypeStr(sb,"class", rowFTField, 29);
                        sb = appendFieldTypeStr(sb,"ignoreCase", rowFTField, 30);
                        sb = appendFieldTypeStr(sb,"words", rowFTField, 31);
                        sb = appendFieldTypeStr(sb,"format", rowFTField, 32);
                        sb.append("/>");
                    }
                }else {
                    sb.append("\n\t<fieldType ");
                    sb = appendFieldTypeStr(sb,"name", rowFTField, 23);
                    sb = appendFieldTypeStr(sb,"class", rowFTField, 24);
                    sb = appendFieldTypeStr(sb,"positionIncrementGap", rowFTField, 25);
                    sb.append(">");
                    // 拼接analyzer
                    sb.append("\n\t\t<analyzer ");
                    sb = appendFieldTypeStr(sb,"type", rowFTField, 26);
                    sb.append(">");
                    // 拼接tokenizer
                    sb.append("\n\t\t\t<tokenizer ");
                    sb = appendFieldTypeStr(sb,"class", rowFTField, 27);
                    sb = appendFieldTypeStr(sb,"mode", rowFTField, 28);
                    sb.append("/>");
                    // 拼接filter
                    sb.append("\n\t\t\t<filter ");
                    sb = appendFieldTypeStr(sb,"class", rowFTField, 29);
                    sb = appendFieldTypeStr(sb,"ignoreCase", rowFTField, 30);
                    sb = appendFieldTypeStr(sb,"words", rowFTField, 31);
                    sb = appendFieldTypeStr(sb,"format", rowFTField, 32);
                    sb.append("/>");
                }
            }else {
                break;
            }
        }
        bw.write(sb.toString());
        bw.flush();
    }


    /**
     * 生成copyField字段
     * @param bw
     * @throws IOException
     */
    private static void createCopyField(BufferedWriter bw) throws IOException {
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
            sb = appendStr(sb, "source", cellCFSource, 19);

            Cell cellCFDest = sheet.getRow(start).getCell(20);
            sb = appendStr(sb, "dest", cellCFDest, 20);

            sb.append("/>");
            bw.write(sb.toString());
            bw.newLine();
        }
    }

    /**
     * 生成uniqueKey字段
     * @param bw
     * @throws IOException
     */
    private static void createUniqueKey(BufferedWriter bw) throws IOException {
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
     * @param bw
     * @throws IOException
     */
    private static void createDynamicField(BufferedWriter bw) throws IOException {
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
            sb = appendStr(sb, "name", cellDynamicField, 9);

            Cell cellDYType = sheet.getRow(start).getCell(10);
            sb = appendStr(sb, "type", cellDYType, 10);

            Cell cellDYIndexed = sheet.getRow(start).getCell(11);
            sb = appendStr(sb, "indexed", cellDYIndexed, 11);

            Cell cellDYStored = sheet.getRow(start).getCell(12);
            sb = appendStr(sb, "stored", cellDYStored, 12);

            Cell cellDYMultiValued = sheet.getRow(start).getCell(13);
            sb = appendStr(sb, "multiValued", cellDYMultiValued, 13);

            sb.append("/>");
            bw.write(sb.toString());
            bw.newLine();
        }
    }

    /**
     * 生成field字段
     * @param bw
     * @throws IOException
     */
    private static void createField(BufferedWriter bw) throws IOException {
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
            sb = appendStr(sb, "name", cellField, 1);

            Cell cellFieldType = sheet.getRow(start).getCell(2);
            sb = appendStr(sb, "type", cellFieldType, 2);

            Cell cellFieldIndexed = sheet.getRow(start).getCell(3);
            sb = appendStr(sb, "indexed", cellFieldIndexed, 3);

            Cell cellFieldStored = sheet.getRow(start).getCell(4);
            sb = appendStr(sb, "stored", cellFieldStored, 4);

            Cell cellFieldRequired = sheet.getRow(start).getCell(5);
            sb = appendStr(sb, "required", cellFieldRequired, 5);

            Cell cellMultiValued = sheet.getRow(start).getCell(6);
            sb = appendStr(sb, "multiValued", cellMultiValued, 6);

            sb.append("/>\n");
            bw.write(sb.toString());
        }
    }

    /**
     * 拼接属性字符串
     * @param sb
     * @param attrName 属性名称
     * @param cell 属性值对应的单元格对象
     * @param cellIndex 列索引
     * @return
     */
    private static StringBuffer appendStr(StringBuffer sb, String attrName, Cell cell, int cellIndex) {
        if(checkCellIsNull(cell)) {
            sb.append(attrName).append("=\"").append(sheet.getRow(6).getCell(cellIndex).getStringCellValue()).append("\" ");
        }else {
            sb.append(attrName).append("=\"").append(cell.getStringCellValue()).append("\" ");
        }
        return sb;
    }

    /**
     * 拼接fieldType属性字符串
     * @param sb
     * @param attrName
     * @param row
     * @param cellIndex
     * @return
     */
    private static StringBuffer appendFieldTypeStr(StringBuffer sb, String attrName, Row row, int cellIndex) {
        if(checkCellIsNull(row.getCell(cellIndex))) {
            Font font = wb.createFont();
            font.setColor(Font.COLOR_RED);
            font.setBold(true);
            CellStyle cellStyle = wb.createCellStyle();
            cellStyle.setFont(font);
            Cell cell = row.getCell(33);
            if(cell == null) {
                cell = row.createCell(33);
                cell.setCellValue(row.getRowNum()+ "," +cellIndex + "字段确定没有值吗?--- ");
            }else {
                cell.setCellValue(cell.getStringCellValue() + "(" + row.getRowNum()+ "," +cellIndex + ")" + sheet.getRow(5).getCell(cellIndex) +"字段单元格确定没有值吗!---");
            }
            cell.setCellStyle(cellStyle);
        }else {
            sb.append(attrName).append("=\"").append(row.getCell(cellIndex).getStringCellValue()).append("\" ");
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
