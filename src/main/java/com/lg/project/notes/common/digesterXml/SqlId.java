package com.lg.project.notes.common.digesterXml;

import java.util.ArrayList;

/**
 * 存储xml中id节点对应的类
 * @author Administrator
 */
public class SqlId {

    //属性value对应的值
    public String value = "";

    //子节点对应的集合变量
    public ArrayList<SqlValue> sql = new ArrayList<>();


    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public void addSqlValue(SqlValue sqlValue){
        this.sql.add(sqlValue);
    }
}
