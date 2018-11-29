package com.xiaoqiang.easyexcel.excel.entity;

import java.util.List;

/**
 * ExcelColumn <br>
 * 〈〉
 *
 * @author XiaoQiang
 * @create 2018-11-29
 * @since 1.0.0
 */
public class ExcelColumn {

    //列名
    private String title;
    //列对应的数据中的field
    private String field;
    //列宽
    private int width=0;
    //子列
    private List<ExcelColumn> children;

    public ExcelColumn(){}

    public ExcelColumn(String title, String field, int width) {
        super();
        this.title = title;
        this.field = field;
        this.width = width;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String getField() {
        return field;
    }

    public void setField(String field) {
        this.field = field;
    }

    public int getWidth() {
        return width;
    }

    public void setWidth(int width) {
        this.width = width;
    }

    public List<ExcelColumn> getChildren() {
        return children;
    }

    public void setChildren(List<ExcelColumn> children) {
        this.children = children;
    }
}
