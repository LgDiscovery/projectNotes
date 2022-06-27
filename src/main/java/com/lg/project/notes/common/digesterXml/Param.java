package com.lg.project.notes.common.digesterXml;

/**
 * 存储xml中param节点对应的类
 * @author Administrator
 */
public class Param {

    public String value = "";

    public String model = "";

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public String getModel() {
        return model;
    }

    public void setModel(String model) {
        this.model = model;
    }
}
