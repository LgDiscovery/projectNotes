package com.lg.project.notes.execlPoi.code.poi;

import com.lg.project.notes.common.model.AjaxResult;
import com.lg.project.notes.common.reflect.ReflectUtils;
import com.lg.project.notes.common.text.Convert;
import com.lg.project.notes.common.utils.DateUtils;
import com.lg.project.notes.common.utils.FileTypeUtils;
import com.lg.project.notes.common.utils.StringUtils;
import com.lg.project.notes.execlPoi.code.annotation.Excel;
import com.lg.project.notes.execlPoi.code.annotation.Excel.ColumnType;
import com.lg.project.notes.execlPoi.code.annotation.Excel.Type;
import com.lg.project.notes.execlPoi.code.annotation.Excels;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
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

    public static final String DOWN_LOAD_PATH = "D:/ruoyi/uploadPath/download";

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

    //******************************** 导出 接口定义地方 start ************************** //


    /**
     * @description: 对list数据源将其里面的数据导入到excel表单
     * @param: response 返回数据
                list 导出数据集合
                sheetName 工作表的名称
     * @return: void
     * @author liuguang
     * @date: 2022/4/3 11:48
     */
    public void exportExcel(HttpServletResponse response, List<T> list, String sheetName){
        exportExcel(response,list,sheetName,StringUtils.EMPTY);
    }

    /**
     * @description: 对list数据源将其里面的数据导入到excel表单
     * @param: response 返回数据
                list 导出数据集合
                sheetName 工作表的名称
                title 标题
     * @return: void
     * @author liuguang
     * @date: 2022/4/3 11:48
     */
    public void exportExcel(HttpServletResponse response, List<T> list, String sheetName,String title){
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setCharacterEncoding("utf-8");
        this.init(list, sheetName, title, Type.EXPORT);
        exportExcel(response);
    }

    /**
     * 对list数据源将其里面的数据导入到excel表单
     *
     * @param list 导出数据集合
     * @param sheetName 工作表的名称
     * @return 结果
     */
    public AjaxResult exportExcel(List<T> list, String sheetName)
    {
        return exportExcel(list, sheetName, StringUtils.EMPTY);
    }

    /**
     * 对list数据源将其里面的数据导入到excel表单
     *
     * @param list 导出数据集合
     * @param sheetName 工作表的名称
     * @param title 标题
     * @return 结果
     */
    public AjaxResult exportExcel(List<T> list, String sheetName, String title)
    {
        this.init(list, sheetName, title, Type.EXPORT);
        return exportExcel();
    }

    /**
     * 对list数据源将其里面的数据导入到excel表单
     *
     * @param sheetName 工作表的名称
     * @return 结果
     */
    public AjaxResult importTemplateExcel(String sheetName)
    {
        return importTemplateExcel(sheetName, StringUtils.EMPTY);
    }

    /**
     * 对list数据源将其里面的数据导入到excel表单
     *
     * @param sheetName 工作表的名称
     * @param title 标题
     * @return 结果
     */
    public AjaxResult importTemplateExcel(String sheetName, String title)
    {
        this.init(null, sheetName, title, Type.IMPORT);
        return exportExcel();
    }

    /**
     * 对list数据源将其里面的数据导入到excel表单
     *
     * @param sheetName 工作表的名称
     * @return 结果
     */
    public void importTemplateExcel(HttpServletResponse response, String sheetName)
    {
        importTemplateExcel(response, sheetName, StringUtils.EMPTY);
    }

    /**
     * 对list数据源将其里面的数据导入到excel表单
     *
     * @param sheetName 工作表的名称
     * @param title 标题
     * @return 结果
     */
    public void importTemplateExcel(HttpServletResponse response, String sheetName, String title)
    {
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setCharacterEncoding("utf-8");
        this.init(null, sheetName, title, Type.IMPORT);
        exportExcel(response);
    }

    //******************************** 导出 接口定义地方 end ************************** //


    //******************************** 导入 接口定义地方 start ************************** //
    /**
     * 对excel表单默认第一个索引名转换成list
     *
     * @param is 输入流
     * @return 转换后集合
     */
    public List<T> importExcel(InputStream is) throws Exception
    {
        return importExcel(is, 0);
    }

    /**
     * 对excel表单默认第一个索引名转换成list
     *
     * @param is 输入流
     * @param titleNum 标题占用行数
     * @return 转换后集合
     */
    public List<T> importExcel(InputStream is, int titleNum) throws Exception
    {
        return importExcel(StringUtils.EMPTY, is, titleNum);
    }

    /**
     * 对excel表单指定表格索引名转换成list
     *
     * @param sheetName 表格索引名
     * @param titleNum 标题占用行数
     * @param is 输入流
     * @return 转换后集合
     */
    public List<T> importExcel(String sheetName, InputStream is, int titleNum) throws Exception
    {
        this.type = Type.IMPORT;
        this.wb = WorkbookFactory.create(is);
        List<T> list = new ArrayList<T>();
        // 如果指定sheet名,则取指定sheet中的内容 否则默认指向第1个sheet
        Sheet sheet = StringUtils.isNotEmpty(sheetName) ? wb.getSheet(sheetName) : wb.getSheetAt(0);
        if (sheet == null)
        {
            throw new IOException("文件sheet不存在");
        }
        boolean isXSSFWorkbook = !(wb instanceof HSSFWorkbook);
        Map<String, PictureData> pictures;
        if (isXSSFWorkbook)
        {
            pictures = getSheetPictures07((XSSFSheet) sheet, (XSSFWorkbook) wb);
        }
        else
        {
            pictures = getSheetPictures03((HSSFSheet) sheet, (HSSFWorkbook) wb);
        }
        // 获取最后一个非空行的行下标，比如总行数为n，则返回的为n-1
        int rows = sheet.getLastRowNum();

        if (rows > 0)
        {
            // 定义一个map用于存放excel列的序号和field.
            Map<String, Integer> cellMap = new HashMap<String, Integer>();
            // 获取表头
            Row heard = sheet.getRow(titleNum);
            for (int i = 0; i < heard.getPhysicalNumberOfCells(); i++)
            {
                Cell cell = heard.getCell(i);
                if (StringUtils.isNotNull(cell))
                {
                    String value = this.getCellValue(heard, i).toString();
                    cellMap.put(value, i);
                }
                else
                {
                    cellMap.put(null, i);
                }
            }
            // 有数据时才处理 得到类的所有field.
            List<Object[]> fields = this.getFields();
            Map<Integer, Object[]> fieldsMap = new HashMap<Integer, Object[]>();
            for (Object[] objects : fields)
            {
                Excel attr = (Excel) objects[1];
                Integer column = cellMap.get(attr.name());
                if (column != null)
                {
                    fieldsMap.put(column, objects);
                }
            }
            for (int i = titleNum + 1; i <= rows; i++)
            {
                // 从第2行开始取数据,默认第一行是表头.
                Row row = sheet.getRow(i);
                // 判断当前行是否是空行
                if (isRowEmpty(row))
                {
                    continue;
                }
                T entity = null;
                for (Map.Entry<Integer, Object[]> entry : fieldsMap.entrySet())
                {
                    Object val = this.getCellValue(row, entry.getKey());

                    // 如果不存在实例则新建.
                    entity = (entity == null ? clazz.newInstance() : entity);
                    // 从map中得到对应列的field.
                    Field field = (Field) entry.getValue()[0];
                    Excel attr = (Excel) entry.getValue()[1];
                    // 取得类型,并根据对象类型设置值.
                    Class<?> fieldType = field.getType();
                    if (String.class == fieldType)
                    {
                        String s = Convert.toStr(val);
                        if (StringUtils.endsWith(s, ".0"))
                        {
                            val = StringUtils.substringBefore(s, ".0");
                        }
                        else
                        {
                            String dateFormat = field.getAnnotation(Excel.class).dateFormat();
                            if (StringUtils.isNotEmpty(dateFormat))
                            {
                                val = DateUtils.parseDateToStr(dateFormat, (Date) val);
                            }
                            else
                            {
                                val = Convert.toStr(val);
                            }
                        }
                    }
                    else if ((Integer.TYPE == fieldType || Integer.class == fieldType) && StringUtils.isNumeric(Convert.toStr(val)))
                    {
                        val = Convert.toInt(val);
                    }
                    else if (Long.TYPE == fieldType || Long.class == fieldType)
                    {
                        val = Convert.toLong(val);
                    }
                    else if (Double.TYPE == fieldType || Double.class == fieldType)
                    {
                        val = Convert.toDouble(val);
                    }
                    else if (Float.TYPE == fieldType || Float.class == fieldType)
                    {
                        val = Convert.toFloat(val);
                    }
                    else if (BigDecimal.class == fieldType)
                    {
                        val = Convert.toBigDecimal(val);
                    }
                    else if (Date.class == fieldType)
                    {
                        if (val instanceof String)
                        {
                            val = DateUtils.parseDate(val);
                        }
                        else if (val instanceof Double)
                        {
                            val = DateUtil.getJavaDate((Double) val);
                        }
                    }
                    else if (Boolean.TYPE == fieldType || Boolean.class == fieldType)
                    {
                        val = Convert.toBool(val, false);
                    }
                    if (StringUtils.isNotNull(fieldType))
                    {
                        String propertyName = field.getName();
                        if (StringUtils.isNotEmpty(attr.targetAttr()))
                        {
                            propertyName = field.getName() + "." + attr.targetAttr();
                        }
                        else if (StringUtils.isNotEmpty(attr.readConverterExp()))
                        {
                            val = reverseByExp(Convert.toStr(val), attr.readConverterExp(), attr.separator());
                        }
                        /*else if (StringUtils.isNotEmpty(attr.dictType()))
                        {
                            val = reverseDictByExp(Convert.toStr(val), attr.dictType(), attr.separator());
                        }*/
                        else if (!attr.handler().equals(ExcelHandlerAdapter.class))
                        {
                            val = dataFormatHandlerAdapter(val, attr);
                        }
                        else if (ColumnType.IMAGE == attr.cellType() && StringUtils.isNotEmpty(pictures))
                        {
                            PictureData image = pictures.get(row.getRowNum() + "_" + entry.getKey());
                            if (image == null)
                            {
                                val = "";
                            }
                            else
                            {
                                byte[] data = image.getData();
                                //val = FileUtils.writeImportBytes(data);
                            }
                        }
                        ReflectUtils.invokeSetter(entity, propertyName, val);
                    }
                }
                list.add(entity);
            }
        }
        return list;
    }
    //******************************** 导入 接口定义地方 end ************************** //

    /**
     * 反向解析值字典值
     *
     * @param dictLabel 字典标签
     * @param dictType 字典类型
     * @param separator 分隔符
     * @return 字典值
     */
   /* public static String reverseDictByExp(String dictLabel, String dictType, String separator)
    {
        return DictUtils.getDictValue(dictType, dictLabel, separator);
    }*/

    /**
     * 反向解析值 男=0,女=1,未知=2
     *
     * @param propertyValue 参数值
     * @param converterExp 翻译注解
     * @param separator 分隔符
     * @return 解析后值
     */
    public static String reverseByExp(String propertyValue, String converterExp, String separator)
    {
        StringBuilder propertyString = new StringBuilder();
        String[] convertSource = converterExp.split(",");
        for (String item : convertSource)
        {
            String[] itemArray = item.split("=");
            if (StringUtils.containsAny(separator, propertyValue))
            {
                for (String value : propertyValue.split(separator))
                {
                    if (itemArray[1].equals(value))
                    {
                        propertyString.append(itemArray[0] + separator);
                        break;
                    }
                }
            }
            else
            {
                if (itemArray[1].equals(propertyValue))
                {
                    return itemArray[0];
                }
            }
        }
        return StringUtils.stripEnd(propertyString.toString(), separator);
    }

    /**
     * 判断是否是空行
     *
     * @param row 判断的行
     * @return
     */
    private boolean isRowEmpty(Row row)
    {
        if (row == null)
        {
            return true;
        }
        for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++)
        {
            Cell cell = row.getCell(i);
            if (cell != null && cell.getCellType() != CellType.BLANK)
            {
                return false;
            }
        }
        return true;
    }

    /**
     * 获取单元格值
     *
     * @param row 获取的行
     * @param column 获取单元格列号
     * @return 单元格值
     */
    public Object getCellValue(Row row, int column)
    {
        if (row == null)
        {
            return row;
        }
        Object val = "";
        try
        {
            Cell cell = row.getCell(column);
            if (StringUtils.isNotNull(cell))
            {
                if (cell.getCellType() == CellType.NUMERIC || cell.getCellType() == CellType.FORMULA)
                {
                    val = cell.getNumericCellValue();
                    if (DateUtil.isCellDateFormatted(cell))
                    {
                        val = DateUtil.getJavaDate((Double) val); // POI Excel 日期格式转换
                    }
                    else
                    {
                        if ((Double) val % 1 != 0)
                        {
                            val = new BigDecimal(val.toString());
                        }
                        else
                        {
                            val = new DecimalFormat("0").format(val);
                        }
                    }
                }
                else if (cell.getCellType() == CellType.STRING)
                {
                    val = cell.getStringCellValue();
                }
                else if (cell.getCellType() == CellType.BOOLEAN)
                {
                    val = cell.getBooleanCellValue();
                }
                else if (cell.getCellType() == CellType.ERROR)
                {
                    val = cell.getErrorCellValue();
                }

            }
        }
        catch (Exception e)
        {
            return val;
        }
        return val;
    }
    /**
     * 对list数据源将其里面的数据导入到excel表单
     *
     * @return 结果
     */
    private AjaxResult exportExcel()
    {
        OutputStream out = null;
        try
        {
            writeSheet();
            String filename = encodingFilename(sheetName);
            out = new FileOutputStream(getAbsoluteFile(filename));
            wb.write(out);
            return AjaxResult.success(filename);
        }
        catch (Exception e)
        {
            log.error("导出Excel异常{}", e.getMessage());
            throw new RuntimeException("导出Excel失败，请联系网站管理员！");
        }
        finally
        {
            IOUtils.closeQuietly(wb);
            IOUtils.closeQuietly(out);
        }
    }
    /**
     * 编码文件名
     */
    public String encodingFilename(String filename)
    {
        filename = UUID.randomUUID().toString() + "_" + filename + ".xlsx";
        return filename;
    }

    /**
     * 获取下载路径
     *
     * @param filename 文件名称
     */
    public String getAbsoluteFile(String filename)
    {
        String downloadPath = ExcelUtil.DOWN_LOAD_PATH + filename;
        File desc = new File(downloadPath);
        if (!desc.getParentFile().exists())
        {
            desc.getParentFile().mkdirs();
        }
        return downloadPath;
    }
    /**
     * @description: 对list数据源将其里面的数据导入到excel表单
     * @param: response
     * @return: void
     * @author liuguang
     * @date: 2022/4/3 11:52
     */
    public void exportExcel(HttpServletResponse response){
        try{
            //1.填充数据sheet数据
            writeSheet();
            //2.把工作薄对象写出去
            wb.write(response.getOutputStream());
        }catch (Exception e){
            log.error("导出Excel异常{}",e.getMessage());
        }finally {
            IOUtils.closeQuietly(wb);
        }
    }

    /**
     * @description: 创建写入数据到Sheet  需要注意一个sheet 和多个sheet处理情况
     * @param:
     * @return: void
     * @author liuguang
     * @date: 2022/4/3 11:57
     */
    public void writeSheet(){
        //1.先判断有多少个sheet
        int sheetNo = Math.max(1,(int)Math.ceil(list.size() * 1.0 / sheetSize));
        for (int index = 0; index < sheetNo; index++) {
            //创建sheet init() 创建过一个 sheet 所以注意逻辑
            createSheet(sheetNo,index);
            // 产生一行
            Row row = sheet.createRow(rownum);
            int column = 0;
            // 写入各个字段的列头名称
            for(Object[] os:fields){
                Excel excel = (Excel)os[1];
                //创建头部标题
                createHeaderCell(excel,row,column++);
            }

            if(Type.EXPORT.equals(type)){
                // 填从数据
                fillExcelData(index,row);
                //在末尾添加统计行
                addStatisticsRow();
            }
        }
    }

    /**
     * @description: 填充excel数据
     * @param: index 序号
                row 单元格行
     * @return: void
     * @author liuguang
     * @date: 2022/4/3 12:36
     */
    public void fillExcelData(int index, Row row) {
        int startNo = index * sheetSize;
        int endNo = Math.min(startNo+sheetSize,list.size());
        for (int i = startNo; i <endNo ; i++) {
            row = sheet.createRow(i+1+rownum -startNo);
            // 得到导出对象.
            T vo =  (T)list.get(i);
            int column = 0;
            for (Object[] os:fields) {
                Field field = (Field) os[0];
                Excel excel = (Excel) os[1];
                //添加数据单元格并设置样式 是否统计处理 处理统计数据
                this.addBodyDataCell(excel, row, vo, field, column++);
            }
        }

    }
    /**
     * @description: 创建统计行
     * @param:
     * @return: void
     * @author liuguang
     * @date: 2022/4/3 15:02
     */
    public void addStatisticsRow()
    {
        if(statistics.size() > 0){
            Row row = sheet.createRow(sheet.getLastRowNum() + 1);
            Set<Integer> keys = statistics.keySet();
            Cell cell = row.createCell(0);
            cell.setCellStyle(styles.get("total"));
            cell.setCellValue("合计");

            for(Integer key:keys){
                cell =row.createCell(key);
                cell.setCellStyle(styles.get("total"));
                cell.setCellValue(DOUBLE_FORMAT.format(statistics.get(key)));
            }
            statistics.clear();
        }
    }

    /**
     * @description: 添加数据单元格 并设置样式
     * @param: attr  注解信息
                row 行
                vo  实体数据
                field 字段信息
                column 序号
     * @return: org.apache.poi.ss.usermodel.Cell
     * @author liuguang
     * @date: 2022/4/3 12:42
     */
    public Cell addBodyDataCell(Excel attr, Row row, T vo, Field field, int column)
    {
        Cell cell = null;
        try{
            //设置行高
            row.setHeight(maxHeight);
            // 根据Excel中设置情况决定是否导出,有些情况需要保持为空,希望用户填写这一列.
            if(attr.isExport()){
                // 创建cell
                cell = row.createCell(column);
                String align = attr.align().value();
                cell.setCellStyle(styles.get("data" + align));
                // 用于读取对象中的属性
                Object value = getTargetValue(vo, field, attr);
                //获取日期格式
                String dateFormat = attr.dateFormat();
                //获取读取内容转表达式 (如: 0=男,1=女,2=未知)
                String readConverterExp = attr.readConverterExp();
                //分隔符，
                String separator = attr.separator();
                //如果是字典类型，请设置字典的type值 (如: sys_user_sex)
                String dictType = attr.dictType();
                //日期格式处理
                if(StringUtils.isNotEmpty(dateFormat) && StringUtils.isNotNull(value)){
                    cell.setCellValue(DateUtils.parseDateToStr(dateFormat,(Date) value));
                    //处理读取内容转表达式
                }else if(StringUtils.isNotEmpty(readConverterExp) && StringUtils.isNotNull(value)){
                    cell.setCellValue(convertByExp(Convert.toStr(value),readConverterExp,separator));
                    // 处理字典类型装换
                }/*else if(StringUtils.isNotEmpty(dictType) && StringUtils.isNotNull(value)){
                    cell.setCellValue(convertDictByExp(Convert.toStr(value), dictType, separator));
                }*/
                // 处理BigDecimal 精度 默认:-1(默认不开启BigDecimal格式化)
                else if(value instanceof BigDecimal && -1 != attr.scale()){
                    cell.setCellValue((((BigDecimal) value).setScale(attr.scale(), attr.roundingMode())).toString());
                }
                //ExcelHandlerAdapter 处理
                else if(!attr.handler().equals(ExcelHandlerAdapter.class)){
                    cell.setCellValue(dataFormatHandlerAdapter(value, attr));
                }else{
                    // 处理除以上其他列类型  NUMERIC(0), STRING(1), IMAGE(2);
                    setCellVo(value, attr, cell);
                }
                // 添加统计数据
                addStatisticsData(column, Convert.toStr(value), attr);
            }

        }catch (Exception e){
            log.error("导出Excel失败{}",e.getMessage());
        }

        return cell;
    }

    /**
     * 合计统计信息
     */
    private void addStatisticsData(Integer index, String text, Excel entity)
    {
        if(Objects.nonNull(entity) && entity.isStatistics()){
            Double temp = 0D;
            if(!statistics.containsKey(index)){
                statistics.put(index,temp);
            }
            try{
                temp = Double.valueOf(text);
            }catch (NumberFormatException e){
                log.error("数据格式错误",e.getMessage());
            }
            statistics.put(index,statistics.get(index)+temp);
        }
    }

    /**
     * 数据处理器
     *
     * @param value 数据值
     * @param excel 数据注解
     * @return
     */
    public String dataFormatHandlerAdapter(Object value, Excel excel)
    {
        try{
            Object instance = excel.handler().newInstance();
            Method method = excel.handler().getMethod("format", new Class[]{Object.class, String[].class});
            value = method.invoke(instance,value,excel.agrs());

        }catch (Exception e){
            log.error("不能格式化数据 " + excel.handler(),e.getMessage());
        }
        return Convert.toStr(value);
    }

    /**
     * 设置其他数据类型单元格信息
     *
     * @param value 单元格值
     * @param attr 注解相关
     * @param cell 单元格信息
     */
    public void setCellVo(Object value, Excel attr, Cell cell)
    {
        if(ColumnType.STRING == attr.cellType()){
            String cellValue = Convert.toStr(value);
            // 对于任何以表达式触发字符 =-+@开头的单元格，直接使用tab字符作为前缀，防止CSV注入。
            if (StringUtils.containsAny(cellValue, FORMULA_STR))
            {
                cellValue = StringUtils.replaceEach(cellValue, FORMULA_STR, new String[] { "\t=", "\t-", "\t+", "\t@" });
            }
            cell.setCellValue(StringUtils.isNull(cellValue) ? attr.defaultValue() : cellValue + attr.suffix());
        }
        else if (ColumnType.NUMERIC == attr.cellType())
        {
            if (StringUtils.isNotNull(value))
            {
                cell.setCellValue(StringUtils.contains(Convert.toStr(value), ".") ? Convert.toDouble(value) : Convert.toInt(value));
            }
        }
        /*else if (ColumnType.IMAGE == attr.cellType())
        {
            ClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, (short) cell.getColumnIndex(), cell.getRow().getRowNum(), (short) (cell.getColumnIndex() + 1), cell.getRow().getRowNum() + 1);
            String imagePath = Convert.toStr(value);
            if (StringUtils.isNotEmpty(imagePath))
            {
                byte[] data = ImageUtils.getImage(imagePath);
                getDrawingPatriarch(cell.getSheet()).createPicture(anchor,
                        cell.getSheet().getWorkbook().addPicture(data, getImageType(data)));
            }
        }*/
    }

    /**
     * 获取画布
     */
    public static Drawing<?> getDrawingPatriarch(Sheet sheet)
    {
        if (sheet.getDrawingPatriarch() == null)
        {
            sheet.createDrawingPatriarch();
        }
        return sheet.getDrawingPatriarch();
    }

    /**
     * 获取图片类型,设置图片插入类型
     */
    public int getImageType(byte[] value)
    {
        String type = FileTypeUtils.getFileExtendName(value);
        if ("JPG".equalsIgnoreCase(type))
        {
            return Workbook.PICTURE_TYPE_JPEG;
        }
        else if ("PNG".equalsIgnoreCase(type))
        {
            return Workbook.PICTURE_TYPE_PNG;
        }
        return Workbook.PICTURE_TYPE_JPEG;
    }

    /**
     * @description: 解析导出值 0=男,1=女,2=未知
     * @param: propertyValue 参数值
                converterExp 翻译注解
                separator 分隔符
     * @return: java.lang.String 解析后值
     * @author liuguang
     * @date: 2022/4/3 13:33
     */
    public static String convertByExp(String propertyValue, String converterExp, String separator)
    {
        StringBuilder propertyString = new StringBuilder();
        String[] convertSource = converterExp.split(",");
        for (String item:convertSource) {
            String[] itemArray = item.split("=");
            if(StringUtils.containsAny(separator,propertyValue)){
                for (String value:propertyValue.split(separator)){
                    if(itemArray[0].equals(value)){
                        propertyString.append(itemArray[1]+separator);
                    }
                }
            }else{
                if(itemArray[0].equals(propertyValue)){
                    return itemArray[1];
                }
            }
        }
        return StringUtils.stripEnd(propertyString.toString(),separator);
    }
    /**
     * @description: 获取bean中的属性值
     * @param: vo 实体对象
                field 字段
                excel 注解
     * @return: java.lang.Object 最终的属性值
     * @author liuguang
     * @date: 2022/4/3 12:53
     */
    private Object getTargetValue(T vo, Field field, Excel excel) throws Exception
    {
        Object o = field.get(vo);
        //处理 属性是其他实体类情况
        /**
         *  @Excels({
         *            @Excel(name = "部门名称", targetAttr = "deptName", type = Excel.Type.EXPORT),
         *            @Excel(name = "部门负责人", targetAttr = "leader", type = Excel.Type.EXPORT)
         *         })
         */
        if(StringUtils.isNotEmpty(excel.targetAttr())){
             String target = excel.targetAttr();
             if(target.contains(".")){
                 String[] targets = target.split("[.]");
                 for (String name:targets) {
                     o = getValue(o,name);
                 }
             }else{
                 o = getValue(o,target);
             }
        }

        return o;
    }

    /**
     * @description: 以类的属性的get方法方法形式获取值
     * @param: o
                name
     * @return: java.lang.Object
     * @author liuguang
     * @date: 2022/4/3 12:58
     */
    private Object getValue(Object o, String name) throws Exception
    {
        if(StringUtils.isNotNull(o) && StringUtils.isNotEmpty(name)){
            Class<?> clazz = o.getClass();
            Field field = clazz.getDeclaredField(name);
            field.setAccessible(true);
            o = field.get(o);
        }
        return o;
    }

    /**
     * @description:  创建工作表
     * @param: sheetNo 多少个sheet 大于1 开始
                index 序号
     * @return: void
     * @author liuguang
     * @date: 2022/4/3 12:08
     */
    public void createSheet(int sheetNo,int index){
        if(sheetNo > 1 && index > 0){
            // 1.根据数据的多少分多个sheet
            this.sheet = wb.createSheet();
            // 2.创建sheet标题
            this.createTitle();
            // 3.给sheet设置名称
            wb.setSheetName(index,sheetName + index);
        }
    }


    /** 
     * @description: 创建头部字段单元格
     * @param: 
     * @return: org.apache.poi.ss.usermodel.Cell
     * @author liuguang
     * @date: 2022/4/2 21:51
     */ 
    public Cell createHeaderCell(Excel attr,Row row,int column){
        Cell cell =row.createCell(column);
        cell.setCellValue(attr.name());
        //实现 excel 注解中prompt提示 和 combo属性
        setDataValidation(attr, row, column);
        //设置头部列的样式
        cell.setCellStyle(styles.get("header"));
        return cell;
    }

    /**
     * @description: excel 注解中prompt提示 和 combo属性 设置 的实现
     * @param: attr
                row
                column
     * @return: void
     * @author liuguang
     * @date: 2022/4/3 11:15
     */
    private void setDataValidation(Excel attr,Row row,int column){
        if(attr.name().indexOf("注：") >=0){
            sheet.setColumnWidth(column,6000);
        }else{
            sheet.setColumnWidth(column,(int)((attr.width()+0.72)*256));
        }
        // 如果设置了提示信息则鼠标放上去提示.
        if(StringUtils.isNotEmpty(attr.prompt())){
            setXSSFPrompt(sheet,"",attr.prompt(),1,100,column,column);
        }
        //如果设置了combo属性则本列只能选择不能输入
        if(attr.combo().length > 0){
            setXSSFCombo(sheet,attr.combo(),1,100,column,column);
        }
    }

    /**
     * 设置某些列的值只能输入预制的数据,显示下拉框.
     *
     * @param sheet 要设置的sheet.
     * @param textlist 下拉框显示的内容
     * @param firstRow 开始行
     * @param endRow 结束行
     * @param firstCol 开始列
     * @param endCol 结束列
     * @return 设置好的sheet.
     */
    public void setXSSFCombo(Sheet sheet, String[] textlist, int firstRow, int endRow, int firstCol, int endCol)
    {
        DataValidationHelper helper = sheet.getDataValidationHelper();
        //加载下拉列表内容
        DataValidationConstraint constraint = helper.createExplicitListConstraint(textlist);
        // 设置数据有效性加载在哪个单元格上,四个参数分别是：起始行、终止行、起始列、终止列
        CellRangeAddressList regions = new CellRangeAddressList(firstRow, endRow, firstCol, endCol);
        // 数据有效性对象
        DataValidation dataValidation = helper.createValidation(constraint, regions);
        // 处理Excel兼容性问题
        if(dataValidation instanceof XSSFDataValidation){
            dataValidation.setSuppressDropDownArrow(true);
            dataValidation.setShowErrorBox(true);
        }else{
            dataValidation.setSuppressDropDownArrow(true);
        }
        sheet.addValidationData(dataValidation);
    }

    /**
     * 设置 POI XSSFSheet 单元格提示
     *
     * @param sheet 表单
     * @param promptTitle 提示标题
     * @param promptContent 提示内容
     * @param firstRow 开始行
     * @param endRow 结束行
     * @param firstCol 开始列
     * @param endCol 结束列
     */
    public void setXSSFPrompt(Sheet sheet, String promptTitle, String promptContent, int firstRow, int endRow,
                              int firstCol, int endCol)
    {
        DataValidationHelper helper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint = helper.createCustomConstraint("DD1");
        CellRangeAddressList regions = new CellRangeAddressList(firstRow,endRow,firstCol,endCol);
        DataValidation dataValidation = helper.createValidation(constraint, regions);
        dataValidation.createPromptBox(promptTitle,promptContent);
        dataValidation.setShowPromptBox(true);
        sheet.addValidationData(dataValidation);
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
        styles.put("dataAuto", style);

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

    /**
     * 获取Excel2003图片
     *
     * @param sheet 当前sheet对象
     * @param workbook 工作簿对象
     * @return Map key:图片单元格索引（1_1）String，value:图片流PictureData
     */
    public static Map<String, PictureData> getSheetPictures03(HSSFSheet sheet, HSSFWorkbook workbook)
    {
        Map<String, PictureData> sheetIndexPicMap = new HashMap<String, PictureData>();
        List<HSSFPictureData> pictures = workbook.getAllPictures();
        if (!pictures.isEmpty())
        {
            for (HSSFShape shape : sheet.getDrawingPatriarch().getChildren())
            {
                HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
                if (shape instanceof HSSFPicture)
                {
                    HSSFPicture pic = (HSSFPicture) shape;
                    int pictureIndex = pic.getPictureIndex() - 1;
                    HSSFPictureData picData = pictures.get(pictureIndex);
                    String picIndex = String.valueOf(anchor.getRow1()) + "_" + String.valueOf(anchor.getCol1());
                    sheetIndexPicMap.put(picIndex, picData);
                }
            }
            return sheetIndexPicMap;
        }
        else
        {
            return sheetIndexPicMap;
        }
    }

    /**
     * 获取Excel2007图片
     *
     * @param sheet 当前sheet对象
     * @param workbook 工作簿对象
     * @return Map key:图片单元格索引（1_1）String，value:图片流PictureData
     */
    public static Map<String, PictureData> getSheetPictures07(XSSFSheet sheet, XSSFWorkbook workbook)
    {
        Map<String, PictureData> sheetIndexPicMap = new HashMap<String, PictureData>();
        for (POIXMLDocumentPart dr : sheet.getRelations())
        {
            if (dr instanceof XSSFDrawing)
            {
                XSSFDrawing drawing = (XSSFDrawing) dr;
                List<XSSFShape> shapes = drawing.getShapes();
                for (XSSFShape shape : shapes)
                {
                    if (shape instanceof XSSFPicture)
                    {
                        XSSFPicture pic = (XSSFPicture) shape;
                        XSSFClientAnchor anchor = pic.getPreferredSize();
                        CTMarker ctMarker = anchor.getFrom();
                        String picIndex = ctMarker.getRow() + "_" + ctMarker.getCol();
                        sheetIndexPicMap.put(picIndex, pic.getPictureData());
                    }
                }
            }
        }
        return sheetIndexPicMap;
    }

}
