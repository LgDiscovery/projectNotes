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
    
    
   
 
 