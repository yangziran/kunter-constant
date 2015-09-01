/**
 * 
 */
package cn.kunter.common.constant.make;

import java.io.FileOutputStream;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import cn.kunter.common.constant.config.PropertyHolder;
import cn.kunter.common.constant.entity.Column;
import cn.kunter.common.constant.entity.Table;

/**
 * 根据表结构生成数据模板
 * @author 阳自然
 * @version 1.0 2015年8月28日
 */
public class MakeExcel {

    public static void main(String[] args) throws Exception {

        List<Table> tables = GetTableConfig.getTableConfig();
        MakeExcel.makerSheet(tables);
    }

    /**
     * 生成Sheet
     * @param tables
     * @throws Exception
     * @author 阳自然
     */
    public static void makerSheet(List<Table> tables) throws Exception {

        // 创建新的Excel 工作簿
        Workbook workbook = new XSSFWorkbook();

        // 生成履历Sheet
        makerHisSheet(workbook);

        // 生成表一览
        makerListSheet(workbook, tables);

        // 遍历表结构创建表设计书
        for (Table table : tables) {
            makerTableSheet(workbook, table);
        }

        // 新建一输出文件流
        FileOutputStream fileOut = new FileOutputStream(PropertyHolder.getConfigProperty("target") + "系统基础数据.xlsx");

        // 把相应的Excel 工作簿存盘
        workbook.write(fileOut);
        fileOut.flush();

        // 操作结束，关闭文件
        fileOut.close();
        workbook.close();
    }

    /**
     * 生成TableSheet
     * @param workbook
     * @param table
     * @throws Exception
     * @author 阳自然
     */
    public static void makerTableSheet(Workbook workbook, Table table) throws Exception {

        // 生成表设计Sheet 备注不规范会出现错误 所以使用TableName为Sheet名称
        Sheet sheet = workbook.createSheet(table.getTableName());
        // 设置默认行高
        sheet.setDefaultRowHeight((short) 350);
        // 设置默认列宽
        sheet.setDefaultColumnWidth(2);
        // 冻结第0列第5行
        sheet.createFreezePane(0, 5);

        CreationHelper createHelper = workbook.getCreationHelper();
        // 列合并
        // 第一行处理
        CellRangeAddress region = new CellRangeAddress(0, 0, 0, 2);
        sheet.addMergedRegion(region);
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0, Cell.CELL_TYPE_STRING);
        cell.setCellStyle(getCellStyleBlueNoBorder(workbook));
        cell.setCellValue("表名");
        region = new CellRangeAddress(0, 0, 3, 12);
        sheet.addMergedRegion(region);
        cell = row.createCell(3, Cell.CELL_TYPE_STRING);
        cell.setCellValue(table.getRemarks());
        region = new CellRangeAddress(0, 0, 13, 15);
        sheet.addMergedRegion(region);
        cell = row.createCell(13, Cell.CELL_TYPE_STRING);
        cell.setCellStyle(getCellStyleBlueNoBorder(workbook));
        cell.setCellValue("物理名");
        region = new CellRangeAddress(0, 0, 16, 25);
        sheet.addMergedRegion(region);
        cell = row.createCell(16, Cell.CELL_TYPE_STRING);
        cell.setCellValue(table.getTableName());
        region = new CellRangeAddress(0, 0, 26, 28);
        sheet.addMergedRegion(region);
        cell = row.createCell(26, Cell.CELL_TYPE_STRING);
        Hyperlink link = createHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);
        link.setAddress("#表一览!A1");
        cell.setCellStyle(getLinkStyleNoBorder(workbook));
        cell.setHyperlink(link);
        cell.setCellValue("返回列表");

        CellStyle sellStyleTitle = getCellStyleBlueNoBorder(workbook);
        sellStyleTitle.setAlignment(CellStyle.ALIGN_CENTER);

        // 第三行处理 标题
        row = sheet.createRow(2);
        for (int j = 0; j < table.getCols().size(); j++) {
            Column column = table.getCols().get(j);

            region = new CellRangeAddress(2, 2, (j * 5) + 5, (j * 5) + 9);
            sheet.addMergedRegion(region);
            cell = row.createCell((j * 5) + 5, Cell.CELL_TYPE_STRING);
            cell.setCellStyle(sellStyleTitle);
            cell.setCellValue(column.getRemarks());
        }

        // 第四行处理 标题
        row = sheet.createRow(3);
        for (int j = 0; j < table.getCols().size(); j++) {
            Column column = table.getCols().get(j);

            region = new CellRangeAddress(3, 3, (j * 5) + 5, (j * 5) + 9);
            sheet.addMergedRegion(region);
            cell = row.createCell((j * 5) + 5, Cell.CELL_TYPE_STRING);
            cell.setCellStyle(sellStyleTitle);
            cell.setCellValue(column.getColumnName());
        }

        // 第五行处理 标题
        row = sheet.createRow(4);
        for (int j = 0; j < table.getCols().size(); j++) {
            Column column = table.getCols().get(j);

            region = new CellRangeAddress(4, 4, (j * 5) + 5, (j * 5) + 9);
            sheet.addMergedRegion(region);
            cell = row.createCell((j * 5) + 5, Cell.CELL_TYPE_STRING);
            cell.setCellStyle(sellStyleTitle);
            cell.setCellValue(column.getLength());
        }

        // 数据行
        for (int i = 0; i < 100; i++) {
            row = sheet.createRow(i + 5);
            region = new CellRangeAddress(i + 5, i + 5, 4, 4);
            sheet.addMergedRegion(region);
            cell = row.createCell(4, Cell.CELL_TYPE_STRING);
            cell.setCellValue(String.valueOf(i + 1));
            for (int j = 0; j < table.getCols().size(); j++) {
                region = new CellRangeAddress(i + 5, i + 5, (j * 5) + 5, (j * 5) + 9);
                sheet.addMergedRegion(region);
                cell = row.createCell((j * 5) + 5, Cell.CELL_TYPE_STRING);
            }
        }
    }

    /**
     * 生成表一览Sheet
     * @param workbook
     * @param tables
     * @throws Exception
     * @author 阳自然
     */
    public static void makerListSheet(Workbook workbook, List<Table> tables) throws Exception {

        // 创建表一览Sheet
        Sheet sheet = workbook.createSheet("表一览");

        // 创建日期格式的样式
        CellStyle cellStyleDate = getCellStyle(workbook);
        DataFormat format = workbook.createDataFormat();
        cellStyleDate.setDataFormat(format.getFormat("yyyy-mm-dd"));
        // 水平居中对齐
        cellStyleDate.setAlignment(CellStyle.ALIGN_CENTER);

        CreationHelper createHelper = workbook.getCreationHelper();

        // 遍历所有的表
        for (int i = 0; i < tables.size(); i++) {
            Table table = tables.get(i);

            // 去除两行列标题
            Row row = sheet.createRow(i + 2);
            // 设置行高
            row.setHeightInPoints(15);

            Cell cell = row.createCell(0, Cell.CELL_TYPE_STRING);
            cell.setCellValue(i + 1);
            cell.setCellStyle(getCellStyle(workbook));
            cell = row.createCell(1, Cell.CELL_TYPE_STRING);
            cell.setCellStyle(getLinkStyle(workbook));
            cell.setCellValue(table.getRemarks());
            Hyperlink link = createHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);
            link.setAddress("#" + table.getTableName() + "!A1");
            cell.setHyperlink(link);
            cell = row.createCell(2, Cell.CELL_TYPE_STRING);
            cell.setCellStyle(getLinkStyle(workbook));
            cell.setCellValue(table.getTableName());
            cell.setHyperlink(link);
            cell = row.createCell(3, Cell.CELL_TYPE_STRING);
            cell.setCellStyle(getCellStyle(workbook));
            cell.setCellValue(table.getRemarks());
            cell = row.createCell(4, Cell.CELL_TYPE_STRING);
            cell.setCellStyle(getCellStyle(workbook));
            cell.setCellValue("自动生成");
            cell = row.createCell(5, Cell.CELL_TYPE_STRING);
            cell.setCellStyle(cellStyleDate);
            cell.setCellValue(new Date());
            cell = row.createCell(6, Cell.CELL_TYPE_STRING);
            cell.setCellStyle(getCellStyle(workbook));
            cell = row.createCell(7, Cell.CELL_TYPE_STRING);
            cell.setCellStyle(cellStyleDate);
        }

        // 在索引0的位置创建行（最顶端的行）
        Row row = sheet.createRow(0);
        // 设置行高
        row.setHeightInPoints(25);

        // 3.1 创建字体，设置其为粗体：
        Font font = getFont(workbook, 12);
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        CellStyle cellStyle = getCellStyle(workbook);
        cellStyle.setFont(font);
        // 水平居中对齐
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);

        // 3.3应用格式
        // 在索引0的位置创建单元格（左上端）
        Cell cell = row.createCell(0, Cell.CELL_TYPE_STRING);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("表一览");

        CellStyle cellStyleBlue = getCellStyleBlue(workbook);
        // 水平居中对齐
        cellStyleBlue.setAlignment(CellStyle.ALIGN_CENTER);

        // 在索引1的位置创建行（第二行）
        row = sheet.createRow(1);
        // 设置行高
        row.setHeightInPoints(30);
        cell = row.createCell(0);
        cell.setCellType(Cell.CELL_TYPE_STRING);
        cell.setCellStyle(cellStyleBlue);
        cell.setCellValue("No.");
        cell = row.createCell(1);
        cell.setCellStyle(cellStyleBlue);
        cell.setCellValue("表名");
        cell = row.createCell(2);
        cell.setCellStyle(cellStyleBlue);
        cell.setCellValue("表物理名");
        cell = row.createCell(3);
        cell.setCellStyle(cellStyleBlue);
        cell.setCellValue("概要");
        cell = row.createCell(4);
        cell.setCellStyle(cellStyleBlue);
        cell.setCellValue("作成者");
        cell = row.createCell(5);
        cell.setCellStyle(cellStyleBlue);
        cell.setCellValue("作成日期");
        cell = row.createCell(6);
        cell.setCellStyle(cellStyleBlue);
        cell.setCellValue("修正者");
        cell = row.createCell(7);
        cell.setCellStyle(cellStyleBlue);
        cell.setCellValue("修正日期");

        // 合并单元格 第一行 A到F
        CellRangeAddress region = new CellRangeAddress(0, 0, 0, 7);
        sheet.addMergedRegion(region);
        // 冻结第1列第2行
        sheet.createFreezePane(8, 2);
        // 设置列宽
        sheet.setColumnWidth(0, 1000);
        sheet.setColumnWidth(1, 6500);
        sheet.setColumnWidth(2, 6500);
        sheet.setColumnWidth(3, 10000);
        sheet.setColumnWidth(4, 2500);
        sheet.setColumnWidth(5, 2500);
        sheet.setColumnWidth(6, 2500);
        sheet.setColumnWidth(7, 2500);
    }

    /**
     * 生成修改履历Sheet
     * @param workbook
     * @throws Exception
     * @author 阳自然
     */
    public static void makerHisSheet(Workbook workbook) throws Exception {

        // 创建修改履历Sheet
        Sheet sheet = workbook.createSheet("修改履历");

        // 日期了类型的单元格格式
        CellStyle cellStyleDate = getCellStyle(workbook);
        DataFormat format = workbook.createDataFormat();
        cellStyleDate.setDataFormat(format.getFormat("yyyy-mm-dd"));
        // 水平居中对齐
        cellStyleDate.setAlignment(CellStyle.ALIGN_CENTER);

        for (int i = 0; i < 33; i++) {
            Row row = sheet.createRow(i);
            // 设置行高
            row.setHeightInPoints(15);
            for (int j = 0; j < 6; j++) {
                Cell cell = row.createCell(j, Cell.CELL_TYPE_STRING);
                cell.setCellStyle(getCellStyle(workbook));
                if (j == 0) {
                    cell.setCellStyle(cellStyleDate);
                }
            }
        }

        // 在索引0的位置创建行（最顶端的行）
        Row row = sheet.createRow(0);
        // 设置行高
        row.setHeightInPoints(25);

        // 设置字体为粗体
        Font font = getFont(workbook, 12);
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        CellStyle cellStyleBold = getCellStyle(workbook);
        cellStyleBold.setFont(font);
        // 水平居中对齐
        cellStyleBold.setAlignment(CellStyle.ALIGN_CENTER);

        // 3.3应用格式
        // 在索引0的位置创建单元格（左上端）
        Cell cell = row.createCell(0, Cell.CELL_TYPE_STRING);
        cell.setCellStyle(cellStyleBold);
        cell.setCellValue("修正履历一览");

        CellStyle cellStyleBlue = getCellStyleBlue(workbook);
        // 水平居中对齐
        cellStyleBlue.setAlignment(CellStyle.ALIGN_CENTER);

        // 在索引1的位置创建行（第二行）
        row = sheet.createRow(1);
        // 设置行高
        row.setHeightInPoints(30);
        cell = row.createCell(0);
        cell.setCellType(Cell.CELL_TYPE_STRING);
        cell.setCellStyle(cellStyleBlue);
        cell.setCellValue("修正日期");
        cell = row.createCell(1);
        cell.setCellStyle(cellStyleBlue);
        cell.setCellValue("修正表名");
        cell = row.createCell(2);
        cell.setCellStyle(cellStyleBlue);
        cell.setCellValue("表物理名");
        cell = row.createCell(3);
        cell.setCellStyle(cellStyleBlue);
        cell.setCellValue("修正内容");
        cell = row.createCell(4);
        cell.setCellStyle(cellStyleBlue);
        cell.setCellValue("修正理由");
        cell = row.createCell(5);
        cell.setCellStyle(cellStyleBlue);
        cell.setCellValue("修正者");

        // 合并单元格 第一行 A到F
        CellRangeAddress region = new CellRangeAddress(0, 0, 0, 5);
        sheet.addMergedRegion(region);
        // 冻结第6列第2行
        sheet.createFreezePane(6, 2);
        // 设置列宽
        sheet.setColumnWidth(0, 3500);
        sheet.setColumnWidth(1, 5500);
        sheet.setColumnWidth(2, 5500);
        sheet.setColumnWidth(3, 10000);
        sheet.setColumnWidth(4, 7000);
        sheet.setColumnWidth(5, 2000);
    }

    /**
     * 普通单元格样式
     * @param workbook
     * @return
     * @author 阳自然
     */
    private static CellStyle getCellStyleNoBorder(Workbook workbook) {

        // 创建格式 普通格式
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFont(getFont(workbook, null));
        // 垂直居中对齐
        cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);

        return cellStyle;
    }

    /**
     * 普通单元格样式
     * @param workbook
     * @return
     * @author 阳自然
     */
    private static CellStyle getCellStyle(Workbook workbook) {

        // 创建格式 普通格式
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFont(getFont(workbook, null));
        // 单元格边框
        cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
        cellStyle.setBorderTop(CellStyle.BORDER_THIN);
        cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
        cellStyle.setBorderRight(CellStyle.BORDER_THIN);
        // 垂直居中对齐
        cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);

        return cellStyle;
    }

    /**
     * 蓝色背景样式 无边框
     * @param workbook
     * @return
     * @author 阳自然
     */
    private static CellStyle getCellStyleBlueNoBorder(Workbook workbook) {

        // 创建格式 普通格式
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFont(getFont(workbook, null));

        // 设置单元格颜色
        cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        cellStyle.setFillForegroundColor(IndexedColors.CORNFLOWER_BLUE.getIndex());

        return cellStyle;
    }

    /**
     * 蓝色背景样式
     * @param workbook
     * @return
     * @author 阳自然
     */
    private static CellStyle getCellStyleBlue(Workbook workbook) {

        CellStyle cellStyle = getCellStyle(workbook);

        // 设置单元格颜色
        cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        cellStyle.setFillForegroundColor(IndexedColors.CORNFLOWER_BLUE.getIndex());

        return cellStyle;
    }

    private static CellStyle getLinkStyleNoBorder(Workbook workbook) {

        Font font = getFont(workbook, null);
        font.setColor(IndexedColors.BLUE.getIndex());

        CellStyle cellStyle = getCellStyleNoBorder(workbook);
        cellStyle.setFont(font);

        return cellStyle;
    }

    private static CellStyle getLinkStyle(Workbook workbook) {

        Font font = getFont(workbook, null);
        font.setColor(IndexedColors.BLUE.getIndex());

        CellStyle cellStyle = getCellStyle(workbook);
        cellStyle.setFont(font);

        return cellStyle;
    }

    /**
     * 获取字体
     * @param workbook
     * @param fontSize
     * @return
     * @author 阳自然
     */
    private static Font getFont(Workbook workbook, Integer fontSize) {

        // 如果字体大小为空 默认设置 10
        if (fontSize == null) {
            fontSize = 10;
        }

        Font font = workbook.createFont();
        font.setFontName("宋体");
        // 设置字体大小
        font.setFontHeightInPoints(fontSize.shortValue());

        return font;
    }
}
