package com.xiaoqiang.easyexcel.excel.entity;

import com.xiaoqiang.easyexcel.excel.constant.ExcelConstant;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

/**
 * PoiStyle <br>
 * 〈样式〉
 *
 * @author XiaoQiang
 * @create 2018-11-29
 * @since 1.0.0
 */
public class PoiStyle {

    /**
     * H样式
     *
     * @param sheet
     * @return
     */

    public static CellStyle hStyle(Sheet sheet) {

        Workbook wb = sheet.getWorkbook();

        CellStyle style = wb.createCellStyle();

        style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 居中

        style.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);// 设置背景色

        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

        // style.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框

        // style.setBorderLeft(HSSFCellStyle.BORDER_MEDIUM);//左边框

        // style.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框

        // style.setBorderRight(HSSFCellStyle.BORDER_MEDIUM);//右边框

        Font font = wb.createFont();

        // font.setFontName("黑体");

        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);//粗体显示

        font.setFontHeightInPoints((short) 14);//设置字体大小

        // HSSFFont font2 = wb.createFont();

        // font2.setFontName("仿宋_GB2312");

        // font2.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);//粗体显示

        // font2.setFontHeightInPoints((short) 12);

        style.setFont(font);//选择需要用到的字体格式

        // style.setWrapText(true);//设置自动换行

        return style;

    }

    /**
     * title 样式
     *
     * @param sheet
     * @return
     */

    public static CellStyle tStyle(Sheet sheet) {

        Workbook wb = sheet.getWorkbook();

        CellStyle style = wb.createCellStyle();

        style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 居中
        Font font = wb.createFont();
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);//粗体显示
        style.setFont(font);//选择需要用到的字体格式

        style.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框

        return style;

    }

    /**
     * 数据列样式
     *
     * @param sheet
     * @return
     */

    public static CellStyle cosStyle(Sheet sheet) {

        Workbook wb = sheet.getWorkbook();

        CellStyle style = wb.createCellStyle();

        style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 居中

        return style;

    }

    /**
     * 数据列样式
     *
     * @param sheet
     * @return
     */

    public static CellStyle cosStyle(Sheet sheet, String cellStyle, String cellFormat) {
        Workbook wb = sheet.getWorkbook();

        CellStyle style = wb.createCellStyle();

        style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 居中

        if ("Date".equals(cellStyle)) {
            if (cellFormat == null || "".equals(cellFormat)) {
                cellFormat = "yyyy-MM-dd";
            }
            style.setDataFormat(wb.createDataFormat().getFormat(cellFormat));//日期格式
        }

        return style;

    }

    /**
     * 数据列样式
     *
     * @param sheet
     * @return
     */

    public static XSSFCellStyle cosStyle(Sheet sheet, String cellStyle, String cellFormat, String color, String cellAlign) {
        Workbook wb = sheet.getWorkbook();

        XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();

        //居左
        if (ExcelConstant.HSSFCELLSTYLE_ALIGN_LEFT.equals(cellAlign)) {
            style.setAlignment(HSSFCellStyle.ALIGN_LEFT); // 居左
        } else if (ExcelConstant.HSSFCELLSTYLE_ALIGN_CENTER.equals(cellAlign)) {
            style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 居中
        } else if (ExcelConstant.HSSFCELLSTYLE_ALIGN_RIGHT.equals(cellAlign)) {
            style.setAlignment(HSSFCellStyle.ALIGN_RIGHT); // 居右
        }
//        else {
//            style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 居中
//        }


        if ("Date".equals(cellStyle)) {
            if (cellFormat == null || "".equals(cellFormat)) {
                cellFormat = "yyyy-MM-dd";
            }
            style.setDataFormat(wb.createDataFormat().getFormat(cellFormat));//日期格式
        }
        if ("Double".equals(cellStyle) && !"".equals(cellFormat)) {
            style.setDataFormat(wb.createDataFormat().getFormat(cellFormat));
        }
        String[] colorArray = color == null | "".equals(color) ? null : color.split(",");
        if (colorArray != null) {
            int red = Integer.parseInt(colorArray[0].trim());
            int green = Integer.parseInt(colorArray[1].trim());
            int blue = Integer.parseInt(colorArray[2].trim());
            XSSFColor xSSFColor = color == null ? null : new XSSFColor(new java.awt.Color(red, green, blue));
            style.setFillPattern(CellStyle.SOLID_FOREGROUND);
            style.setFillForegroundColor(xSSFColor);
        } else {
//            XSSFColor xSSFColor = new XSSFColor(new java.awt.Color(255, 255, 255));
//            style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//            style.setFillForegroundColor(xSSFColor);
        }
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框
        return style;

    }

    public static CellStyle getRedBorderStyle(Sheet sheet) {
        Workbook wb = sheet.getWorkbook();
        CellStyle style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        // 设置单元格字体
		/*Font headerFont = workbook.createFont();
		headerFont.setFontHeightInPoints((short)14);
		headerFont.setColor(HSSFColor.RED.index);
		headerFont.setFontName("宋体");
		style.setFont(headerFont);
		style.setWrapText(true);*/

        // 设置单元格边框及颜色
        style.setLeftBorderColor(HSSFColor.RED.index);
        style.setTopBorderColor(HSSFColor.RED.index);
        style.setRightBorderColor(HSSFColor.RED.index);
        style.setBottomBorderColor(HSSFColor.RED.index);
        style.setBorderBottom((short) 1);
        style.setBorderLeft((short) 1);
        style.setBorderRight((short) 1);
        style.setBorderTop((short) 1);
        style.setWrapText(true);
        return style;
    }
}
