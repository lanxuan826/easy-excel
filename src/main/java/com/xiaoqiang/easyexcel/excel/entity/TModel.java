package com.xiaoqiang.easyexcel.excel.entity;

import com.xiaoqiang.easyexcel.excel.annotation.Export;
import com.xiaoqiang.easyexcel.excel.constant.ExcelConstant;
import com.xiaoqiang.easyexcel.excel.util.DateUtil;

import java.util.Date;

/**
 * TModel <br>
 * 〈〉
 *
 * @author XiaoQiang
 * @create 2018-11-29
 * @since 1.0.0
 */
public class TModel {

    @Export
    private boolean boo;

    @Export
    private Integer ints;

    @Export(colName = "类型",width = 5000)
    private String type;

    @Export(colName = "名称",cellColor="242, 220,219")
    private String name;

    @Export
    private  boolean bs;

    @Export(colName = "creation_date",width = 5000,cellStyle="Date",cellFormat= DateUtil.DATE_TIME_FORMAT_DEFAULT,cellAlign= ExcelConstant.HSSFCELLSTYLE_ALIGN_RIGHT)
    private  String creation_date;

    @Export(colName = "creDate",width = 5000,cellStyle = "Date",cellFormat = DateUtil.DATE_TIME_FORMAT_DEFAULT,cellAlign = ExcelConstant.HSSFCELLSTYLE_ALIGN_CENTER)
    private Date creDate;

    @Export(cellStyle="Double" ,cellFormat="#,##0.00")
    private String doubleValue;

    @Export(cellStyle="Integer")
    private String intValue;


    public boolean isBoo() {
        return boo;
    }

    public void setBoo(boolean boo) {
        this.boo = boo;
    }

    public Integer getInts() {
        return ints;
    }

    public void setInts(Integer ints) {
        this.ints = ints;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public boolean isBs() {
        return bs;
    }

    public void setBs(boolean bs) {
        this.bs = bs;
    }

    public String getCreation_date() {
        return creation_date;
    }

    public void setCreation_date(String creation_date) {
        this.creation_date = creation_date;
    }

    public Date getCreDate() {
        return creDate;
    }

    public void setCreDate(Date creDate) {
        this.creDate = creDate;
    }

    public String getDoubleValue() {
        return doubleValue;
    }

    public void setDoubleValue(String doubleValue) {
        this.doubleValue = doubleValue;
    }

    public String getIntValue() {
        return intValue;
    }

    public void setIntValue(String intValue) {
        this.intValue = intValue;
    }
}
