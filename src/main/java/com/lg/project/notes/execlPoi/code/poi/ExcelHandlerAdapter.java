package com.lg.project.notes.execlPoi.code.poi;

/**
 * @ClassName ExcelHandlerAdapter
 * @Description Excel数据格式处理适配器
 * @Author liuguang
 * @Date 2022/4/2 17:13
 * @Version 1.0
 */
public interface ExcelHandlerAdapter {

    /**
     * @description: 格式化
     * @param: value 单元格数据值 args excel注解args参数组
     * @return: java.lang.Object
     * @author liuguang
     * @date: 2022/4/2 18:18
     */
    Object format(Object value,String[] args);
}
