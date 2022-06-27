package com.lg.project.notes.common.digesterXml;

/**
 * 提供处理时间和参数的有关方法
 * @author Administrator
 */
public class DateAndParameter {

    public static String join(String[] value,String split){
        if (value!=null){
            int length = value.length;
            switch (length){
                case 0: return "";
                case 1: return value[0];
                default:
                    StringBuffer buffer = new StringBuffer();
                    for (int i=0;i<length;i++){
                        buffer.append(value[i]).append(split);
                    }
                    return buffer.substring(0,buffer.length()-1);
            }
        }else {
            return "";
        }
    }

    public static String join(String[] value){
        return join(value,",");
    }

    /**
     *  将 ？用值替换
     * @param sentence
     * @param value
     * @return
     */
    public static String replace(String sentence,String value){
        if (value == null || value.equals("")){
            return "";
        }
        StringBuffer sql = new StringBuffer(sentence);
        int location = sql.indexOf("?");
        while (location !=-1){
            if (value.indexOf(",") == -1){
                sql.replace(location,location+1,value);
            }else{
                int equalLocation = sql.indexOf("=");
                boolean flagIn = false;
                if (equalLocation <0){
                    flagIn = true;
                    equalLocation = sql.indexOf("?");
                }
                String value_1 = value;
                if (sql.charAt(++location) == '\''){
                    location++;
                    value_1 = "'" +value.replaceAll(",","','")+"'";
                }
                String[] totalArray = value_1.split(",");
                if (totalArray.length<=1000){
                    if (flagIn){
                        sql.replace(equalLocation,location,value_1);
                    }else {
                        sql.replace(equalLocation,location,"in ("+value_1+")");
                    }
                }else{
                    StringBuffer sbResult = new StringBuffer("(");
                    int num = totalArray.length/1000;
                    for (int i=0;i<num+1;i++){
                        StringBuffer tempSql = new StringBuffer();
                        if (i==num){
                            String[] tempArray = new String[totalArray.length-i*1000];
                            System.arraycopy(totalArray,i*1000,tempArray,0,totalArray.length-i*1000);
                            if (flagIn){
                                sbResult.append(new StringBuffer(tempSql.replace(equalLocation,location,join(tempArray).toString())));
                            }else {
                                sbResult.append(new StringBuffer(tempSql.replace(equalLocation,location,"in ("+join(tempArray)+")").toString()));
                            }
                        }else {
                            String[] tempArray = new String[1000];
                            System.arraycopy(totalArray,i*1000,tempArray,0,1000);
                            if (flagIn){
                                sbResult.append(new StringBuffer(tempSql.replace(equalLocation,location,join(tempArray).toString())));
                            }else {
                                sbResult.append(new StringBuffer(tempSql.replace(equalLocation,location,"in ("+join(tempArray)+")").toString()));
                            }
                            sbResult.append(" or ");
                        }
                    }
                    sbResult.append(")");
                    sql = sbResult;
                }
            }
            location = sql.indexOf("?",location+1);
        }
        return sql.toString();

    }
}
