package com.lg.project.notes.common.digesterXml;

import java.util.ArrayList;

/**
 * 存储xml中sql节点对应的类
 * @author Administrator
 */
public class SqlValue {

    public String value = "";

    public ArrayList<Param> param = new ArrayList<>();

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public void addParam(Param param){
        this.param.add(param);
    }
}
