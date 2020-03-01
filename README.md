# Easy-Excel

## 它是什么?

easy-excel 是基于 Apache POI 框架的一款扩展封装库，让我们在开发中更快速的完成导入导出的需求。 尽管很多人会提出 poi 能干这事儿为什么还要封装一层呢？

easy-excel 很大程度上简化了代码、让使用者更轻松的 读、写 Excel 文档，也不用去关心格式兼容等问题，很多时候我们在代码中会写很多的 for 循环，各种 getXXXIndex 来获取行或列让代码变的更臃肿。多个项目之间打一枪换一个地方，代码 Copy 来 Copy 去十分凌乱， 如果你也在开发中遇到类似的问题，那么 easy-excel 是你值得一试的工具。 

## 特性

**支持:**

- 基于 Java 8 开发
- 简洁的 API 操作
- 注解驱动
- 可配置列顺序
- 自定义导入/导出转换器
- 导出支持自定义样式处理器
- 模板导出
- 支持 Excel 2003、2007格式

**不支持将来可能支持:**

- 多 sheet 导入/导出


## 快速开始

### 引入依赖

>  还没有

### 导出

下面是我们的 Java 模型类，用于存储 Excel 的行数据。

**第一步:** 为我们的 pojo 添加 `@ExcelColumn` 注解，后面会给出各个参数的意思。

```java
@Data
@AllArgsConstructor
public class Product {

    @ExcelColumn(columnName = "id", index = 10)
    private Integer id;

    @ExcelColumn(columnName = "价格", index = 20, centToYuan = true, suffix = " 元", columnNameCellStyleHandler =
            YellowBgCellStyleHandler.class)
    private Long price;

    @ExcelColumn(columnName = "创建日期", index = 60, columnNameCellStyleHandler = GreyBgCellStyleHandler.class)
    private OffsetDateTime created;

    @ExcelColumn(columnName = "名称", index = 40)
    private String name;

    @ExcelColumn(columnName = "是否是新品", index = 41, trueToStr = "新品", falseToStr = "非新品")
    private Boolean isNew;

    @ExcelColumn(columnName = "订单状态", index = 50, enumKey = "name", prefix = "状态: ", cellStyleHandler =
            RedFontCellStyleHandler.class)
    private StateEnum stateEnum;

    @ExcelColumn(columnName = "状态变更日期", index = 55)
    private LocalDateTime updateTime;

    @ExcelColumn(columnName = "备注", index = 70)
    private String other;

}

```

**第二步:** 调用 `ExcelWriter` 类进行导出操作, write() 方法 会返回一个 `Workbook `对象

```java
@Test
public void testExport() throws IOException {
    Workbook workbook = ExcelWriter.create(ExcelType.XLS)
            .sheetName("商品数据")
            .sheetHeader("--2月份商品数据--")
            .write(products, Product.class);

    File file = new File(productFile);
    OutputStream outputStream = new FileOutputStream(file);
    workbook.write(outputStream);
    outputStream.close();
}
```

导出表格:

![3DWpTO.png](https://s2.ax1x.com/2020/02/28/3DWpTO.png)

## 导入

假设我们现在有这样一张 excel 表格：

![3D4YnI.png](https://s2.ax1x.com/2020/02/28/3D4YnI.png)

**第一步:** 给需要接收excel数据的 pojo 实体类添加注解: `@ExcelColumn`

```java
@Data
@AllArgsConstructor
@NoArgsConstructor
public class Employee {

    @ExcelColumn(index = 10, columnName = "ID")
    private Integer id;
    @ExcelColumn(index = 20, columnName = "名称")
    private String name;
    @ExcelColumn(index = 30, columnName = "工资", yuanToCent = true, suffix = " 元")
    private Long salary;
    @ExcelColumn(index = 40, columnName = "性别",enumKey = "name")
    private Gender gender;
    @ExcelColumn(index = 50, columnName = "生日")
    private OffsetDateTime birthday;
}
```

**需要注意的是,导入操作,接受的实体类必须有无参构造器,否则不能出入成功**

**第二步:** 调用 `ExcelReader` 类的方法

```java
@Test
public void test_import_xlsx() throws FileNotFoundException {
    InputStream fileInputStream = new FileInputStream(employeeFile_Xlsx);
    List<Employee> employees = ExcelReader
            .read(fileInputStream, ExcelType.XLSX)
            .columnNameRowNum(1)
            .toPojo(Employee.class);
    Assert.assertNotNull(employees);
}
```

# @ExcelColumn 注解参数

`@ExcelColumn`注解提供了常用的 转换操作,包括**列名定义,列排序,Boolean值转换,前缀后缀,日期格式化,分转元,元转分** 等等.

对于更加复杂的转换,提供了转换器接口:

* **IWriteConverter**: 导出转换器
* **IReadConverter**: 导入转换器

对于需要自定义样式的需求,分别提供了 对于 `columnName` 和 `rowData` 的样式处理器:

* **ICellStyleHandler**: 样式处理器

| 参数                       | 解释                                                         |
| -------------------------- | ------------------------------------------------------------ |
| columnName                 | xcel 每一列的名称                                            |
| index                      | 排序(仅对导出有效),支持不连续的整数                          |
| datePattern                | 日期格式,默认: yyyy/MM/dd,只支持 OffsetDateTime和 LocalDateTime |
| enumKey                    | 枚举导入使用的key,序列化枚举使用                             |
| trueToStr                  | true 转换字符串(仅对导出有效)                                |
| strToTrue                  | 字符串转 true(仅对导入有效)                                  |
| falseToStr                 | false 转换(仅对导出有效)                                     |
| strToFalse                 | 字符串 转 false(仅对导入有效)                                |
| prefix                     | 前缀                                                         |
| suffix                     | 后缀                                                         |
| centToYuan                 | 是否启用 分转元(仅对导出有效) ,仅支持 Integer 和 Long 类型   |
| yuanToCent                 | 是否启用 元转分(仅对导入有效),仅支持 Integer 和 Long 类型    |
| writeConverter             | 导出转换器                                                   |
| readConverter              | 导入转换器                                                   |
| cellStyleHandler           | rowData 样式处理器                                           |
| columnNameCellStyleHandler | columnName 样式处理器                                        |

# API

查看 [wiki](https://github.com/itguang/easy-excel/wiki)

# 测试
本项目采用 github action CI 自动构建,每当 master 分支有变动时就会触发构建,运行所有单元测试。

目前的 jacoco 测试报告：

![3cQy1H.png](https://s2.ax1x.com/2020/03/01/3cQy1H.png)

> 导入导出的核心代码基本覆盖到

# Thanks

* [Apache POI](https://poi.apache.org/)

# License

[Apache2.0](https://github.com/itguang/easy-excel/blob/master/LICENSE)



