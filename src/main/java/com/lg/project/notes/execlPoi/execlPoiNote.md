## 自定义execl导入导出功能的实现 

### Step01 引入maven依赖

    <!-- excel工具 -->
         <dependency>
             <groupId>org.apache.poi</groupId>
             <artifactId>poi-ooxml</artifactId>
             <version>4.1.2</version>
          </dependency>

### Step02 实现逻辑

    1.自定义一个 @Execl 注解 和 @Execls 注解 

    2.定义两个实体类 SysUser,SysDept 用来测试使用
    
    3.扩展定义一个Excel数据格式处理适配器ExcelHandlerAdapter在注解中使用处理一些特殊的数据
    
    4.在需要导出的实体类属性字段中使用注解Excel定义的各种方法定义导出模式
    
    5.添加ExcelUtil工具类实现自定义注解的excel多功能的导出
    
### Step03 ExcelUtil 主要方法的简介

    1.
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

    2.
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
    
    3.
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
   
 
 