package com.lg.project.notes.execlPoi.code.poi;

import com.lg.project.notes.execlPoi.code.annotation.Excel;
import com.lg.project.notes.execlPoi.code.annotation.Excel.Type;
import com.lg.project.notes.execlPoi.code.annotation.Excels;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.text.DecimalFormat;
import java.util.*;
import java.util.stream.Collectors;

/**
 * @ClassName ExcelUtil
 * @Description Excel相关处理
 * @Author liuguang
 * @Date 2022/4/2 19:34
 * @Version 1.0
 */
public class ExcelUtil<T> {
    private static final Logger log = LoggerFactory.getLogger(ExcelUtil.class);

    public static final String[] FORMULA_STR = {"=","-","+","@"};

    /**
     * Excel sheet最大行数，默认65536
     */
    public static final int sheetSize = 65536;

    /**
     * 工作表名称
     */
    private String sheetName;

    /**
     * 导出类型（EXPORT:导出数据 IMPORT:导入模板）
     */
    private Type type;

    /**
     * 工作薄对象
     */
    private Workbook wb;

    /**
     * 工作表对象
     */
    private Sheet sheet;

    /**
     * 样式列表  用map 封装起来不同样式
     */
    private Map<String, CellStyle> styles;

    /**
     * 导入导出数据
     */
    private List<T> list;

    /**
     * 注解列表
     */
    private List<Object[]> fields;


    /**
     * 当前行号
     */
    private int rownum;

    /**
     * 标题
     */
    private String title;

    /**
     * 最大高度
     */
    private short maxHeight;

    /**
     * 统计列表
     */
    private Map<Integer,Double> statistics = new HashMap<>();

    /**
     * 数字格式
     */
    private static final DecimalFormat DOUBLE_FORMAT = new DecimalFormat("######0.00");

    /**
     * 实体对象
     */
    public Class<T> clazz;

    /**
     * @description: 通过构造方法注入需要导出的数据实体 根据实体获取Excel注解信息实现高逼格的导出功能
     * @param: clazz 实体类
     * @return:
     * @author liuguang
     * @date: 2022/4/2 19:51
     */
    public ExcelUtil(Class<T> clazz){
        this.clazz = clazz;
    }


    /**
     * @description: 初始化数据
     * @param:  list  导入导出的数据
                sheetName 工作表名称
                title 标题
                type  导出类型
     * @return: void
     * @author liuguang
     * @date: 2022/4/2 20:06
     */
    public void init(List<T> list,String sheetName,String title,Type type){
        //1.判断数据是否为空 为空初始化一个对象保证不抛出空指针异常
        if(Objects.isNull(list)){
            list = new ArrayList<>();
        }
        //2.否则初始化相关定的属性数据
        this.list = list;
        this.sheetName = sheetName;
        this.title = title;
        this.type = type;
        //3.得到所有定义字段 初始化 fields 保存字段和注解信息 并实现注解的自定义排序
        createExcelField();
        //4.创建工作簿 初始化 workBook sheet 和 各种列表样式styles
        createWorkBook();
        //5.创建excel第一行标题 多sheet 要考虑rownum的初始化设置逻辑
        createTitle();
    }

    /** 
     * @description: 创建头部字段单元格
     * @param: 
     * @return: org.apache.poi.ss.usermodel.Cell
     * @author liuguang
     * @date: 2022/4/2 21:51
     */ 
    public Cell createCell(Excel attr,Row row,int column){
        Cell cell =row.createCell(column);
        cell.setCellValue(attr.name());
        //setDataValidation(attr, row, column); TODO
        cell.setCellStyle(styles.get("header"));
        return cell;
    }
    
     /**
      * @description: 创建excel第一行标题
      * @param:
      * @return: void
      * @author liuguang
      * @date: 2022/4/2 21:26
      */
    private void createTitle(){
        //1.判断是否有标题需要设置
        if(StringUtils.isNotEmpty(title)){
            /**
             * 这个很关键 rownum++ 一个sheet rownum 自增 从1开始
             * 多个sheet 现初始化 rownum = 0  在自增 使得 多个sheet 都是从 1开始
             */
            Row titleRow = sheet.createRow(rownum == 0? rownum++:0);
            titleRow.setHeightInPoints(30); //设置行高
            Cell titleCell =titleRow.createCell(0); // 创建单元格
            titleCell.setCellStyle(styles.get("title")); //设置标题单元格样式
            titleCell.setCellValue(title); //设置单元格值
            //标题一般比较长 所以一般操作是合并单元格
            sheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(),
                    titleRow.getRowNum(),titleRow.getRowNum(),this.fields.size() -1));

        }
    }
    /**
     * @description: 创建工作簿 初始化 workBook sheet 和 各种列表样式styles
     * @param:
     * @return: void
     * @author liuguang
     * @date: 2022/4/2 21:06
     */
    private void createWorkBook(){
        this.wb = new SXSSFWorkbook(500);
        this.sheet = wb.createSheet();
        //可能数据过多 分sheet导出 初始化的时候设置第一个sheet
        wb.setSheetName(0,sheetName);
        //获取单元格各种样式
        this.styles = createStyles(wb);
    }

    /**
     * @description: 获取单元格各种样式
     * @param: wb
     * @return: java.util.Map<java.lang.String,org.apache.poi.ss.usermodel.CellStyle>
     * @author liuguang
     * @date: 2022/4/2 21:11
     */
    private Map<String,CellStyle> createStyles(Workbook wb){
        // 写入各条记录,每条记录对应excel表中的一行
        Map<String, CellStyle> styles = new HashMap<String, CellStyle>();

        //1. title 标题样式 创建一个新的单元格样式并将其添加到工作簿的样式表中
        CellStyle style = wb.createCellStyle();
        //设置单元格的水平对齐类型
        style.setAlignment(HorizontalAlignment.CENTER);
        //设置单元格的垂直对齐类型
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        //创建一个新字体并将其添加到工作簿的字体表中
        Font titleFont = wb.createFont();
        titleFont.setFontName("Arial"); //设置字体的名称（即 Arial）
        titleFont.setFontHeightInPoints((short)16); //设置字体高度
        titleFont.setBold(true);
        style.setFont(titleFont);
        styles.put("title",style);

        //2.设置数据样式
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        Font dataFont = wb.createFont();
        dataFont.setFontName("Arial");
        dataFont.setFontHeightInPoints((short) 10);
        style.setFont(dataFont);
        styles.put("data", style);

        //3.设置头部样式 header
        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font headerFont = wb.createFont();
        headerFont.setFontName("Arial");
        headerFont.setFontHeightInPoints((short) 10);
        headerFont.setBold(true);
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(headerFont);
        styles.put("header", style);

        //4.设置统计样式 total
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        Font totalFont = wb.createFont();
        totalFont.setFontName("Arial");
        totalFont.setFontHeightInPoints((short) 10);
        style.setFont(totalFont);
        styles.put("total", style);

        // 5.根据注解 Align align() default Align.AUTO 设置字段对齐方式（0：默认；1：靠左；2：居中；3：靠右）
        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
        style.setAlignment(HorizontalAlignment.LEFT);
        styles.put("dataLeft", style);

        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
        style.setAlignment(HorizontalAlignment.CENTER);
        styles.put("dataCenter", style);

        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
        style.setAlignment(HorizontalAlignment.RIGHT);
        styles.put("dataRight", style);
        return styles;

    }
    /**
     * @description: 获取实体类添加了Excel注解字段和注解信息保存在fields中 并实现注解中按照自定义的排序导出excel
     * @param:
     * @return: void
     * @author liuguang
     * @date: 2022/4/2 20:15
     */
    private void createExcelField(){
        //1.获取字段信息
        this.fields = getFields();
        //2.根据注解定义的排序对导出字段进行排序
        this.fields = fields.stream().
                sorted(Comparator.comparing(objects -> ((Excel)objects[1]).sort())).
                collect(Collectors.toList());
        //3.根据注解获取最大行高
        this.maxHeight = getRowHeight();
    }

    /**
     * @description: 根据注解获取最大行高
     * @param:
     * @return: short
     * @author liuguang
     * @date: 2022/4/2 20:37
     */
    private short getRowHeight(){
        double maxHeight = 0;
        for (Object[] os:this.fields){
            Excel excel = (Excel)os[1];
            maxHeight = Math.max(maxHeight,excel.height());
        }
        return (short) (maxHeight*20);
    }

    /**
     * @description: 获取字段注解信息
     * @param:
     * @return: java.util.List<java.lang.Object[]{Field,EXCEL}>
     * @author liuguang
     * @date: 2022/4/2 20:22
     */
    private List<Object[]> getFields(){

        //1.定义返回结果集合
        List<Object[]> fields = new ArrayList<>();

        //2.定义一个字段集合
        List<Field> tempFields = new ArrayList<>();

        //3.获取这个实体类中父类声明的字段属性并加入到tempFields
        tempFields.addAll(Arrays.asList(clazz.getSuperclass().getDeclaredFields()));

        //4.获取这个实体类中的声明的字段并加入到tempFields
        tempFields.addAll(Arrays.asList(clazz.getDeclaredFields()));

        //5.遍历tempFields 取出添加了Excel和Excels注解的字段 并将字段和注解信息放入fields返回
        for (Field field:tempFields) {
            //1.单注解
            if(field.isAnnotationPresent(Excel.class)){
                Excel attr = field.getAnnotation(Excel.class);
                if(Objects.nonNull(attr) && (attr.type() == Type.ALL || attr.type() == type)){
                    field.setAccessible(true);
                    fields.add(new Object[]{field,attr});
                }
            }
            //2.多注解
            if(field.isAnnotationPresent(Excels.class)){
                Excels attrs = field.getAnnotation(Excels.class);
                Excel[] excels = attrs.value();
                for (Excel attr:excels){
                    if(Objects.nonNull(attr) && (attr.type() == Type.ALL || attr.type() == type)){
                        field.setAccessible(true);
                        fields.add(new Object[]{field,attr});
                    }
                }
            }
        }
        return fields;
    }

}
