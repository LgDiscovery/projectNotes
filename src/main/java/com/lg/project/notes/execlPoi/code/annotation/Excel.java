package com.lg.project.notes.execlPoi.code.annotation;

import com.lg.project.notes.execlPoi.code.poi.ExcelHandlerAdapter;
import org.apache.commons.compress.archivers.dump.DumpArchiveEntry;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
import java.math.BigDecimal;

/**
 * 自定义导出Execl数据注解
 * @author liuguang
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface Excel {

    /**
     * 导出时在execl中排序 按照该属性排序 导出字段
     * @return
     */
    int sort() default Integer.MAX_VALUE;

    /**
     * 导出到execl中的名字
     * @return
     */
    String name() default "";

    /**
     * 日期时间格式 如：yyyy-MM-dd
     * @return
     */
    String dateFormat() default "";

    /**
     * 如果是字典类型，请设置字典的type值 (如: sys_user_sex)
     */
    String dictType() default "";

    /**
     * 读取内容转表达式 (如: 0=男,1=女,2=未知)
     * @return
     */
    String readConverterExp() default "";

    /**
     * 分隔符，读取字符串组内容
     */
    String separator() default ",";

    /**
     * BigDecimal 精度 默认:-1(默认不开启BigDecimal格式化)
     */
    int scale() default -1;

    /**
     * BigDecimal 舍入规则 默认:BigDecimal.ROUND_HALF_EVEN
     */
    int roundingMode() default BigDecimal.ROUND_HALF_EVEN;

    /**
     * 导出类型（0数字 1字符串）
     * @return
     */
    ColumnType cellType() default ColumnType.STRING;

    /**
     * 导出时在excel中每个列的高度 单位为字符
     * @return
     */
    double height() default 14;

    /**
     * 导出时在excel中每个列的宽 单位为字符
     */
    double width() default 16;

    /**
     * 文字后缀,如% 90 变成90%
     * @return
     */
    String suffix() default "";

    /**
     * 当值为空时，字段的默认值
     * @return
     */
    String defaultValue() default "";

    /**
     * 提示信息
     * @return
     */
    String prompt() default "";

    /**
     * 设置只能选择不能输入的列内容
     * @return
     */
    String[] combo() default {};

    /**
     * 是否导出数据，有些需求要我们导出一份模板，这是标题需要但是内容需要用户手工填写
     * @return
     */
    boolean isExport() default true;

    /**
     * 另一个类中的属性名称,支持多级获取,以小数点隔开
     * @return
     */
    String targetAttr() default "";

    /**
     * 是否自动统计数据,在最后追加一行统计数据总和
     * @return
     */
    boolean isStatistics() default false;

    /**
     * 导出字段对齐方式（0：默认；1：靠左；2：居中；3：靠右）
     * @return
     */
    Align align() default Align.AUTO;

    /**
     * 自定义数据处理器
     * @return
     */
    Class<?> handler() default ExcelHandlerAdapter.class;

    /**
     * 自定义数据处理器参数
     */
    String[] agrs() default {};

    /**
     * 导出类型（0：导出导入；1：仅导出；2：仅导入）
     * @return
     */
    Type type() default Type.ALL;

    enum Type{

        ALL(0),EXPORT(1),IMPORT(2);

         private final int value;

         Type(int value){
             this.value = value;
         }

         public int value(){
             return this.value;
         }
    }

    enum Align{

        AUTO(0),LEFT(1),CENTER(2),RIGHT(3);

        private final int value;

        Align(int value){
            this.value = value;
        }

        public int value(){
            return this.value;
        }
    }

    enum ColumnType{
        NUMERIC(0),STRING(1),IMAGE(2);
        private final int value;

        ColumnType(int value){
            this.value = value;
        }

        public int value(){
            return this.value;
        }
    }
}
