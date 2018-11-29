package com.xiaoqiang.easyexcel.excel.annotation;

import java.lang.annotation.*;

/**
 * 导出annotation <br>
 * 〈导出注解〉
 *
 * @author XiaoQiang
 * @create 2018-11-28
 * @since 1.0.0
 */
@Target({ElementType.FIELD,ElementType.METHOD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Export {

    /**
     * 列名称
     * @return
     */
    String colName() default "col_";


    /**
     * 单元格宽度
     * @return
     */
    int width() default  2500;

    /**
     * 单元格value,默认“”
     * @return
     */
    String value() default  "";

    /**
     * 单元格类型
     * @return
     */
    String cellStyle() default "String";

    /**
     * 单元格格式
    *单元格的日期格式,数字格式
     * 数字格式可以设置format，设置数字显示的样式
     * 例如：
     * cellStyle="Double",cellFormat="#,##0.0000"
     * 显示为带有千分位的数值
     * cellFormat的格式为#,##0.0000，小数点后面的0的个数代表显示数值的小数位数
     * @return
     */
    String cellFormat() default "";

    /**
     * 单元格颜色
     * @return
     */
    String cellColor() default  "";

    /**
     * 单元格对齐方式 默认居中
     * @return
     */
    String cellAlign() default "HSSFCellStyle.ALIGN_CENTER";

}
