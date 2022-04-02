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
    
    3.定义一个Excel数据格式处理适配器ExcelHandlerAdapter在注解中使用处理一些特殊的数据
    

    
   
 
 