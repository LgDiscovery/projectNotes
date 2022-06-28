package com.lg.project.notes.common.digesterXml;


import org.apache.commons.digester3.Digester;
import org.xml.sax.SAXException;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Objects;

/**
 * xml 文件中读取sql模板与传入参数拼接成sql语句
 * @author Administrator
 */
public class MakeSql {

    public static final String XML_URL = "queryStatSQL.xml";

    private static MakeSql makeSql = null;

    public static HashMap<String,SqlId> idMap = new HashMap();

    private MakeSql(){};

    private static MakeSql init() {
        MakeSql makeSql = new MakeSql();
        Digester digester = new Digester();
        digester.push(makeSql);
        //根据xml的结构将数据存入对象
        digester.addObjectCreate("map/id",SqlId.class);
        digester.addSetProperties("map/id","value","value");
        digester.addSetNext("map/id","addId");
        digester.addObjectCreate("map/id/sql",SqlValue.class);
        digester.addSetProperties("map/id/sql","value","value");
        digester.addSetNext("map/id/sql","addSqlValue");
        digester.addObjectCreate("map/id/sql/param",Param.class);
        digester.addSetProperties("map/id/sql/param","value","value");
        digester.addSetNext("map/id/sql/param","addParam");
        try {
            digester.parse(new File(MakeSql.class.getClassLoader().getResource(XML_URL).getPath().replaceAll("%20"," ")));
        } catch (IOException e) {
            e.printStackTrace();
        } catch (SAXException e) {
            e.printStackTrace();
        }
        return (MakeSql) digester.getRoot();
    }

    public static MakeSql getInstance(){
        if(MakeSql.idMap.size() == 0){
            MakeSql.makeSql = MakeSql.init();
        }
        return MakeSql.makeSql;
    }

    public void addId(SqlId sqlId){
        MakeSql.idMap.put(sqlId.value,sqlId);
    }

    public String getSql(String id){
        return getSql(id,0);
    }

    public String getSql(String id ,int index){
        SqlValue sqlValue = getSqlValue(id, index);
        return sqlValue == null ? "":sqlValue.value;
    }

    private SqlValue getSqlValue(String id,int index){
        SqlId sqlId = idMap.get(id);
        return sqlId == null ? null : sqlId.sql.get(index);
    }

    /**
     *
     * @param key  xml中配置的id的value属性值
     * @param index id 中多个sql标签的下标值
     * @param form 查询参数
     * @param startPosition 查询参数和开始用参数替换？的起始位置（多为sql中存在不希望转换 ？ 的情况）
     * @return 拼接好的带查询参数的sql语句
     */
    public String getSQLFormat(String key,int index,Object form ,int startPosition){
        //参数防止sql注入
        String re ="select|update|and|or|delete|truncate|char|into|substr|ascii|declare|exec|count|master|drop|execute|table";
        String[] reArray = re.split("\\|");
        SqlValue sqlValue = getSqlValue(key, index);
        if(Objects.nonNull(sqlValue)){
            ArrayList<Param> paramList = sqlValue.param;
            if(paramList !=null && paramList.size() > 0){
                Object[] obj = new Object[paramList.size()];
                Class<?> formClass = form.getClass();
                Object value = null;
                Param param = null;
                Method method = null;
                try {
                    for (int i=0,j=paramList.size();i<j;i++){
                        param = paramList.get(i);
                        formClass.getMethod("get"
                                +param.value.substring(0,1).toUpperCase()
                                +param.value.substring(1),null);
                        value = method.invoke(form,null); // 参数的值
                        //防止参数sql注入
                        String tempV = value.toString().toLowerCase();
                        for (int k = 0;k<reArray.length;k++){
                            if(tempV.indexOf(reArray[k]) >= 0){
                                if("or".equals(reArray[k]) && tempV.indexOf("'") <0){
                                    continue;
                                }
                                value = "";
                            }
                        }
                         if(value instanceof String[]){
                                value = DateAndParameter.join((String[]) value);
                         }
                         if("".equals(param.model)){
                             obj[i]=value;
                         }else{
                             value = DateAndParameter.replace(param.model,(String) value);
                             obj[i] = value.equals("") ? "" : "AND" + value;
                         }
                    }
                    return getSQL(key,index,obj,startPosition);
                }catch (Exception e){
                    e.printStackTrace();
                }
            }else{
                return sqlValue.value;
            }
        }
        return "";
    }

    public String getSQL(String key,int index,Object[] obj,int startPosition){
        SqlValue sqlValue = getSqlValue(key, index);
        if (sqlValue != null){
            String sql = sqlValue.value;
            if(!"".equals(sql)){
                int toIndex = -1;
                for (int i=1;i<startPosition;i++){
                    toIndex = sql.indexOf('?',toIndex+1);
                }
                StringBuffer buffer = new StringBuffer(sql.length() * 2);
                buffer.append(sql);
                for (int i=0,j=obj.length;i<j;i++){
                    toIndex = buffer.indexOf("?",toIndex +1);
                    if (obj[i]== null){
                        buffer.replace(toIndex,toIndex+1,"");
                    }else {
                        buffer.replace(toIndex,toIndex+1,obj[i].toString());
                    }
                }
                sql = buffer.toString();
                return sql.replaceAll("(where|WHERE)\\s*(and|AND)","WHERE");
            }
        }
        return "";
    }

    public String getSQL(String key,Object[] obj,int startPosition){
        return getSQL(key,0,obj,startPosition);
    }

}
