package com.xiaoqiang.easyexcel.excel.util;

import com.xiaoqiang.easyexcel.excel.annotation.Export;
import com.xiaoqiang.easyexcel.excel.constant.ExcelConstant;
import com.xiaoqiang.easyexcel.excel.entity.PoiStyle;
import com.xiaoqiang.easyexcel.excel.entity.TModel;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.StringUtils;

import java.io.File;
import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.*;

/**
 * Poi工具类<br>
 * 〈poi导出excel〉
 *
 * @author XiaoQiang
 * @create 2018-11-28
 * @since 1.0.0
 */
public class PoiUtil {

    public final static Logger logger = LoggerFactory.getLogger(PoiUtil.class);

    /**
     * 判断文件夹是否存在
     *
     * @param filepath
     * @return
     */

    private static boolean isExistsFile(String filepath) {
        File file = new File(filepath);
        if (file.exists()) {
            return true;
        }
        return false;
    }


    /**
     * 创建文件夹
     *
     * @param filePath
     */
    private static void mkFile(String filePath) {
        File file = new File(filePath);
        file.mkdirs();
    }


    /**
     * 组装Excel列表的title
     *
     * @param clazz
     * @param <T>
     * @return
     */
    private static <T> String getTitle(Class<T> clazz) {

        StringBuffer buffer = new StringBuffer();

        String str = null;

        int i = 1;

        try {
            Field[] fields = clazz.getDeclaredFields();
            for (Field field : fields) {

                if ("boolean".equals(field.getType().getSimpleName())) {
                    str = "is" + field.getName().substring(0, 1).toUpperCase() + field.getName().substring(1);
                } else {
                    str = "get" + field.getName().substring(0, 1).toUpperCase() + field.getName().substring(1);
                }

                if (field.isAnnotationPresent(Export.class)) {
                    Export export = field.getAnnotation(Export.class);
                    buffer.append("=").append("col_".equals(export.colName()) ? export.colName()
                            + i : export.colName());

                }

                i++;
            }

            if (buffer.length() <= 0) {

                throw new Exception("error : PoiUtil.getTitle title is null");

            } else {

                return buffer.toString().substring(1);

            }


        } catch (Exception e) {

            e.printStackTrace();

        }

        return buffer.toString();


    }


    /**
     * 设置标题
     *
     * @param sheet
     * @param h     大标题
     * @param title 列名称
     * @return int row
     */

    private static int writeTitle(Sheet sheet, String h, String title, Object... obj) {

        int row = 0;

        if (title != null && title.length() > 0) {

            String[] t = title.split("=");

            if (h != null && h.length() > 0) {

                Row hrow = sheet.createRow(row); // 大标题行

                Cell hcell = hrow.createCell(0);

                hcell.setCellValue(h);

                sheet.addMergedRegion(new CellRangeAddress(row, row, 0, t.length > 0 ? t.length - 1 : 0));// 合并标题

                hcell.setCellStyle(PoiStyle.hStyle(sheet));// 头部样式

                row++;

            }

            if (t.length > 0) {

                Row hssfrow = sheet.createRow(row);// sheet.createRow(row); //
                // 标题

                row++;

                for (int i = 0; i < t.length; i++) {

                    Cell cell = hssfrow.createCell(i);

                    cell.setCellValue(t[i]);

                    cell.setCellStyle(PoiStyle.tStyle(sheet));

                }

            }

        }

        return row;

    }


    /**
     * POI 导出EXCEL
     *
     * @param clazz 需要导出的Model
     * @param list  需要导出的List(类型必须为clazz)
     * @param obj   obj[0]==文件名字<br>
     * @return map.get(" filename ")=文件名称 map.get("filePath")=文件全路径
     */

    public static <T> Map<String, String> exportExcel(Class<T> clazz, List<T> list, Object... obj) {
        String filePath = ExcelConstant.EXCEL_PATH + "/expFile";
        Map<String, String> map = new HashMap<String, String>();
        FileOutputStream out = null;
        try {
            if (!isExistsFile(filePath)) {
                mkFile(filePath);
            }
            SXSSFWorkbook workbook = new SXSSFWorkbook(999999);
            Sheet sheet = workbook.createSheet();// 默认的sheet0
            String title = getTitle(clazz); // excelTitle
            String fileName = obj == null ? "" : obj.length <= 0 ? "" : obj[0] + "";
            int bRow = writeTitle(sheet, fileName, title, obj); // 数据列表开始行

            Field[] fields = clazz.getDeclaredFields();
            int colNum = 0; //列序号
            for (Field field :fields) {
                String fieldName = field.getName();
                if ("boolean".equals(field.getType().getSimpleName())) {
                    fieldName = "is" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
                } else {
                    fieldName = "get" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
                }
                if (field.isAnnotationPresent(Export.class)) {
                    Method method = clazz.getDeclaredMethod(fieldName);
                    Export exp = field.getAnnotation(Export.class);
                    String expCellStyle = exp.cellStyle();
                    String cellFormat = exp.cellFormat();
                    String color = exp.cellColor();
                    String cellAlign = exp.cellAlign();
                    CellStyle cellStyle = PoiStyle.cosStyle(sheet, expCellStyle, cellFormat, color, cellAlign);
                    int rowNum = 0; //行序号
                    for (T t : list) {
                        Row row;
                        if (colNum == 0) {
                            row = sheet.createRow((rowNum++) + bRow);
                        } else {
                            row = sheet.getRow((rowNum++) + bRow);
                        }
                        Cell cell = row.createCell(colNum);
                        Object robj = method.invoke(t);
                        if (robj == null) {
                            robj = exp.value();
                        }cell.setCellStyle(cellStyle);
                        if ("Date".equals(expCellStyle)) {
                            cell.setCellValue(DateUtil.objToDate(robj, cellFormat));
                        } else if ("Integer".equals(expCellStyle)) {
                            cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
                            if (StringUtils.isEmpty(robj)) {
                                robj = "0";
                            }
                            cell.setCellValue(Integer.parseInt(robj.toString()));
                        } else if ("Double".equals(expCellStyle)) {
                            cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
                            if (StringUtils.isEmpty(robj)) {
                                robj = "0";
                            }
                            cell.setCellValue(Double.parseDouble(robj.toString()));
                        } else {
                            cell.setCellValue(robj.toString());
                        }
                    }
                    sheet.setColumnWidth(colNum++, exp.width());
                }
            }
            map.put("filename", fileName == "" ? System.currentTimeMillis() + ".xlsx" : java.net.URLEncoder.encode(fileName + ".xlsx", "UTF8"));
            if (filePath.endsWith("/")) {
                filePath = filePath + System.currentTimeMillis() + ".xlsx";
            } else {
                filePath = filePath + "/" + System.currentTimeMillis() + ".xlsx";
            }
            map.put("filePath", filePath);
            out = new FileOutputStream(new File(filePath));
            workbook.write(out);
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("error:    PoiUtil.expExcel " + e.getMessage());
        }
        return map;
    }





    public static void main(String[] args) {

        List<TModel> list = new ArrayList<TModel>();

        for (int i = 0; i < 10; i++) {

            TModel a = new TModel();
            a.setType("类型_" + (i + 1));
            a.setName("名字_" + (i + 1));
            a.setCreation_date("2015-07-09 12:22:22");
            a.setIntValue("398");
            a.setCreDate(new Date());
            list.add(a);
        }
        // System.out.println(getTitle(TModel.class) );
        System.out.println(exportExcel(TModel.class, list, new Object[]{"a"}));
    }
}
