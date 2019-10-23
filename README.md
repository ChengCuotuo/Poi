## 使用 poi 完成 excel 文件的导入导出

*实现的内容：前端点击导入根据选择的 excel 文件向数据库中存储信息，点击导出将数据库中所有用户的信息导出*

[ github 仓库地址]("https://github.com/ChengCuotuo/Poi"  "Poi") 持续更新中

数据库文件：

``` mysql
DROP TABLE IF EXISTS `student`;
CREATE TABLE `student` (
  `id` int(11) NOT NULL COMMENT '用户 id',
  `name` varchar(50) NOT NULL COMMENT '用户名',
  `gender` tinyint(4) DEFAULT NULL COMMENT '性别',
  `bir` datetime DEFAULT NULL COMMENT '生日',
  `tel` varchar(20) DEFAULT NULL COMMENT '手机号',
  `email` varchar(320) DEFAULT NULL COMMENT '邮箱',
  `address` varchar(100) DEFAULT NULL COMMENT '居住地',
  `major` varchar(100) DEFAULT NULL COMMENT '专业',
  `school` varchar(100) DEFAULT NULL COMMENT '所在院校',
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
```

在 maven 中使用到的依赖：

``` tex
	<parent>
        <groupId>org.springframework.boot</groupId>
        <artifactId>spring-boot-starter-parent</artifactId>
        <version>2.0.4.RELEASE</version>
        <relativePath/>
    </parent>

        <!-- https://mvnrepository.com/artifact/org.apache.poi/poi -->
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi</artifactId>
            <version>4.0.1</version>
        </dependency>

       <!-- swagger 接口文档 -->
        <!-- https://mvnrepository.com/artifact/io.springfox/springfox-swagger2 -->
        <dependency>
            <groupId>io.springfox</groupId>
            <artifactId>springfox-swagger2</artifactId>
            <version>2.9.2</version>
        </dependency>

        <!-- https://mvnrepository.com/artifact/io.springfox/springfox-swagger-ui -->
        <dependency>
            <groupId>io.springfox</groupId>
            <artifactId>springfox-swagger-ui</artifactId>
            <version>2.9.2</version>
        </dependency>

```

主要的：使用 springboot 2.0.4，poi 4.0.1，swagger 2.9.2

*swagger 低版本的测试导出的时候文件格式会出错*

使用 mybatis.generator 自动生成实体类和 mapper 内容

``` tex
<generatorConfiguration>

    <!-- 本地数据库驱动程序jar包的全路径 -->
    <classPathEntry location="E:\repository\mysql\mysql-connector-java\5.7.21"/>

    <context id="context" targetRuntime="MyBatis3">
        <commentGenerator>
            <property name="suppressDate" value="true"/>
            <property name="suppressAllComments" value="true" />
            <!--其中suppressDate是去掉生成日期那行注释，suppressAllComments是去掉所有的注-->
        </commentGenerator>

        <!-- 数据库的相关配置 -->
        <jdbcConnection driverClass="com.mysql.jdbc.Driver" connectionURL="jdbc:mysql://localhost:3306/school" userId="root" password="root"/>

        <javaTypeResolver>
            <property name="forceBigDecimals" value="false"/>
        </javaTypeResolver>

        <!-- 实体类生成的位置 -->
        <javaModelGenerator targetPackage="cn.nianzuochen.reportform.dao" targetProject="src/main/java">
            <property name="enableSubPackages" value="false"/>
            <property name="trimStrings" value="true"/>
        </javaModelGenerator>

        <!-- *Mapper.xml 文件的位置 -->
        <sqlMapGenerator targetPackage="cn.nianzuochen.reportform.mapper" targetProject="src/main/java">
            <property name="enableSubPackages" value="false"/>
        </sqlMapGenerator>

        <!-- Mapper 接口文件的位置 -->
        <javaClientGenerator targetPackage="cn.nianzuochen.reportform.mapper" targetProject="src/main/java" type="XMLMAPPER">
            <property name="enableSubPackages" value="false"/>
        </javaClientGenerator>

        <table tableName="student" domainObjectName="Student" enableCountByExample="false" enableDeleteByExample="false" enableSelectByExample="false" enableUpdateByExample="false"/>

    </context>
</generatorConfiguration>
```

## 一、导出文件

1. 创建文件
2. 设置文件样式
3. 填充数据
4. 返回文件



### 1.创建文件

​	首先，了解一下一个 excel 的结构，一个 excel 文件就是一个 workbook ，在 workbook 中包含多个 sheet，每个 sheet 就是看到的表格文件了，由多行多列组成。

> Excel 文件
>
> > workbook(HSSFWorkbook)
> >
> > > sheet (HSSFSheet)
> > >
> > > > row (HSSFRow)
> > > >
> > > > > cell (HSSFCell)

``` java
HSSFWorkbook workbook = new HSSFWorkbook();
HSSFSheet sheet = workbook.createSheet(sheetName);
```



### 2.设置文件样式

``` java
/**
     * ToDo 添加表头的样式
     * 设置表头
     * @param sheet 目标 sheet
     * @param headName 表头名称
     * @return
     */
    public void setHead(HSSFSheet sheet, LinkedHashSet<String> head) {
        headName = new LinkedHashSet<>(head);
        int size = headName.size();
        // 设置默认的宽度
        for (int i = 0; i < size; i ++) {
            sheet.setColumnWidth(i, 10 * 256);
        }
        // 设置表头
        HSSFRow headRow = sheet.createRow(0);
        int i = 0;
        for (String s : head) {
            HSSFCell cell = headRow.createCell(i++);
            // 写入内容
            cell.setCellValue(s);
        }
    }
```

### 3. 填充数据

``` java
/**
* 仅传递 sheet 和 实体类信息，根据实体类的 get 方法
* @param sheet 
* @param infos
* @param <T>
* @return
*/
public <T> ResponseEntity<byte[]> addInfo(HSSFSheet sheet, List<T> infos) {
        return  addInfo(sheet, null, infos);
    }
    /**
     * 将数据信息插入到表格中，根据传递的实体类的 get 方法
     * @param sheet
     * @param headToColumn key 表示表头显示的信息，value 则是对应插入进去的属性值名称，并据此调用 get 方法获取属性值
     * @param infos
     * @return
     */
    public <T> ResponseEntity<byte[]> addInfo(HSSFSheet sheet, Map<String, String> headToColumn, List<T> infos) {
        HttpHeaders headers = null;
        ByteArrayOutputStream byteArrayOutputStream = null;
        /**
         * 如果 headName 为空则用此处的 headToColumn 的 keyList 去创建表头
         * 如果 headToColumn 的 keyList 值和指定的 headName 不匹配将用此处的 headToColumn 进行更新
         * 如果 headName 和 headToColumn 均为空，则使用 infos 的中对象的属性名称集合来创建表头（通过 get 方法），创建 headToColumn
         * 然后，根据 headToColumn 中的键值对添加信息到 sheet
         */

        if (headName == null) {
            if (headToColumn == null) {
                // 根据传递的对象的 get 方法设置表头
                // 保证顺序，使用 HashMap 比较合理
                headToColumn = new LinkedHashMap<>();
                // 根据第一个对象进行设置表头
                T first = infos.get(0);
                // 获取该类的所有 get 方法名称
                // 获取当前类
                Class clazz = first.getClass();
                // 获取类中的所有方法
                Method[] methods = clazz.getDeclaredMethods();
                for (Method method : methods) {
                    String methodName = method.getName();
                    // 根据类的 get 方法名称确定类中的属性
                    if (methodName.indexOf("get") >= 0) {
                        methodName = methodName.substring(3);
//                        System.out.println("methodName: " + methodName);
                        headToColumn.put(methodName, methodName);
                    }
                }
                headName = new LinkedHashSet<>(headToColumn.keySet());
                // 重新创建 sheet
                setHead(sheet, headName);
            } else {
                // 直接使用传递的 headToColumn 的 keySet 赋值给 headName
                headName = new LinkedHashSet<>(headToColumn.keySet());
                // 重新创建 sheet
                setHead(sheet, headName);
            }
        } else {
            if (headToColumn != null) {
                // 判断两个是否相等
                boolean isSame = true;
                LinkedHashSet<String> keys = new LinkedHashSet<>(headToColumn.keySet());
                // 判断两个集合是否相同
                Iterator iterator1 = keys.iterator();
                Iterator iterator2 = headName.iterator();
                while (iterator1.hasNext()) {
                    if (!iterator1.next().equals(iterator2.next())) {
                        isSame = false;
                        break;
                    }
                }
                // 不相等要使用 headToColumn 的 keySet 进行赋值
                if (!isSame) {
                    headName = new LinkedHashSet<>(headToColumn.keySet());
                    setHead(sheet, headName);
                }
            } else {
                // 直接使用 headName 中的值赋值 headColumn
                for (String s : headName) {
                    headToColumn.put(s, s);
                }
            }
        }
        // 根据 headToColumn 向 sheet 中添加数据
        for (int i = 0; i < infos.size(); i++) {
            // 为每一个对象创建一行 Row
            HSSFRow row = sheet.createRow(i + 1);
            T t = infos.get(i);
            Class clazz = t.getClass();
            Set<String> keys = headToColumn.keySet();
            int index = 0;
            for (String key: keys) {
                // 将每个属性的 get 方法获取属性值添加到每个 Cell 中
                try {
                    // 根据属性名获取方法
                    Method method = clazz.getMethod("get" + headToColumn.get(key));
                    // 获取方法执行的结果
                    Object obj = method.invoke(t);
                    String returnTypeName = method.getReturnType().getSimpleName();
                    // 所有的信息都看作 String 类型添加
                    if (obj != null) {
//                        System.out.println("Index: " + index + ", ObjInfo: " + obj.toString());
                        // 处理事件类型，进行格式化
                        if ("Date".equals(returnTypeName)) {
                            String value = dateFormat.format((Date)obj);
                            row.createCell(index).setCellValue(value);
                        } else {
                            row.createCell(index).setCellValue(obj.toString());
                        }

                    }
                }  catch (NoSuchMethodException e) {
                    e.printStackTrace();
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                } catch (InvocationTargetException e) {
                    e.printStackTrace();
                }
                index++;
            }
        }
//        System.out.println(fileName);
        try {
            headers = new HttpHeaders();
            headers.setContentDispositionFormData("attachment",
                    new String((fileName + ".xls").getBytes("UTF-8"), "iso-8859-1"));
            headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
            byteArrayOutputStream = new ByteArrayOutputStream();
            workbook.write(byteArrayOutputStream);
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return new ResponseEntity<byte[]>(Objects.requireNonNull(byteArrayOutputStream).toByteArray(),
                headers, HttpStatus.CREATED);
    }
```

另一种插入方式

``` java
	/**
 	* ToDo 增加一些验证
     * Key 表示表头名称，value 则是实际的数据，相当于给出了每一列的数据
     * @param sheet
     * @param stuinfos 需要保证每个 List 中信息条数相同
     * @return
     */
    public <T> ResponseEntity<byte[]> addInfo(HSSFSheet sheet, Map<String, List<T>> infos) {
        HttpHeaders headers = null;
        ByteArrayOutputStream byteArrayOutputStream = null;
        // 如果当前的 headName 和 head 不相同的时候，用 head 重新设置表头
        boolean isSame = true;
        LinkedHashSet<String> keys = (LinkedHashSet<String>)infos.keySet();
        if (headName != null) {
            // 判断两个集合是否相同
            Iterator iterator1 = keys.iterator();
            Iterator iterator2 = headName.iterator();
            while (iterator1.hasNext()) {
                if (!iterator1.next().equals(iterator2.next())) {
                    isSame = false;
                    break;
                }
            }
            // 不相等要使用 headToColumn 的 keySet 进行赋值
            if (!isSame) {
                headName = (LinkedHashSet<String>)infos.keySet();
                setHead(sheet, headName);
            }
        } else {
            headName = (LinkedHashSet<String>)infos.keySet();
            setHead(sheet, headName);
        }
        // 将数据添加到表格中
        List<Object> rowInfo = null;
        List<String> heads = new ArrayList<>(headName);
        // 获取由多少行
        int rowNum = infos.get(heads.get(0)).size();
        for (int i = 0; i < rowNum; i++) {
            HSSFRow row = sheet.createRow(i + 1);
            rowInfo = new ArrayList<>();
            // 产生一行信息，并添加到 excel 中
            int index = 0;
            for (String head : headName) {
                // 获取 cell 值
                String cellValue = infos.get(head).get(i).toString();
                row.createCell(index++).setCellValue(cellValue);
            }
        }
        try {
            headers = new HttpHeaders();
            headers.setContentDispositionFormData("attachment",
                    new String((fileName + ".xls").getBytes("UTF-8"), "iso-8859-1"));
            headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
            byteArrayOutputStream = new ByteArrayOutputStream();
            workbook.write(byteArrayOutputStream);
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return new ResponseEntity<byte[]>(Objects.requireNonNull(byteArrayOutputStream).toByteArray(),
                headers, HttpStatus.CREATED);
    }
```

*使用 swagger 进行测试*

## 二、导入文件

*需要的数据是：实体类（包含 get 和 set 方法）、可以构成改实体类的 excel 文件、excel 表头和实体类属性对应的 Map*

*使用 Java 的反射*

1. 统计指定的实体类的属性，按照 get 方法确定
2. 提供的 Map 的 key 值（指定的 excel 标题）要包含于从 excel 中提取出来的标题
3. 提供的 Map 的 value 值（指定的实体类属性）要包含于从实体中提取出的属性
4. 符合上述条件后，生成实体对象，调用实体类的 set 方法将 excel 中提供的属性添加到实体中
5. 返回根据 excel 生成的实体集合

``` java
public class ExtractExcel2Object <T> {
    // 参考对象
    private Class<T> basic;
    private Set<String> attrMethods = new HashSet<>();;
    private final SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");

    public ExtractExcel2Object(Class<T> basic) {
        this.basic = basic;
        initGetMethods();
    }

    /**
     * 提取提供的类的所有 get 方法对应的属性
     */
    private void initGetMethods ()  {
        try {
            T entity = basic.newInstance();
            Class clazz = entity.getClass();
            // 获取提供的类的所有方法
            Method[] methods = clazz.getMethods();
            for (Method m : methods) {
                // 获取方法名称
                String methodName = m.getName();
                // 提取 get 方法
                if (methodName.indexOf("get") >= 0) {
                    attrMethods.add(methodName.substring(3));
                }
            }
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (InstantiationException e) {
            e.printStackTrace();
        }
    }

    /**
     * 未提供 excel 表头和基础类的对应关系，将按照基础类提取出来的属性名进行类提取
     * @param file
     * @return
     */
    public List<T> extract(MultipartFile file) {
        Map<String, String> head2Attr = new HashMap<>();
        for (String attr : attrMethods) {
            head2Attr.put(attr, attr);
        }
        return extract(head2Attr, file);
    }

    /**
     * ToDo 编写异常处理类提醒各种返回为 null 的各种情况
     * 根据提供的表头和基础类属性对应的 map 创建实体类
     * @param head2Attr 提供的表头和属性对应的关系，key 是表头，value 是属性获取方法
     * @param file excel 问价
     * @return
     */
    public List<T> extract(Map<String, String> head2Attr, MultipartFile file) {
        /**
         * head2Attr 的表头信息 key 值，要包含于 file 的表头
         * head2Attr 的 value 值，要包含与当前的 attrMethods
         * 否则都会出现信息无法匹配
         * 信息匹配后，需要对 head2Attr 进行重排，例如： 1 attrName
         * 最后，用 excel 表格信息生成实体类集合
         */
        // 存储生成的实体类
        List<T> entities = new ArrayList<>();
        Set<String> keys = head2Attr.keySet();
        ArrayList<String> values = new ArrayList<>(head2Attr.values());
        // head2Attr 的 value 值，要包含与当前的 attrMethods
        for (String value : values) {
            // 提供的属性值不存在提供的基础类中
            if (!attrMethods.contains(value)) {
                return null;
            }
        }

        // 提取 excel 的标题，仅处理单 sheet 的 excel 表格
        try {
            Map<Integer, String> column2Method = new LinkedHashMap<>();
            // 存储 excel 标题名称
            List<String> headNames = new ArrayList<>();
            HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(file.getInputStream()));
            // ToDo 待处理sheet 大于 1
            // int sheetNum = workbook.getNumberOfSheets();
            HSSFSheet sheet = workbook.getSheetAt(0);
            // 获取第一行即标题行
            HSSFRow row = sheet.getRow(0);
            // 获取列数，即多少个 cell
            int cellNum = row.getLastCellNum();
            // 获取 String  类型标题
            for (int i = 0; i < cellNum; i++) {
                HSSFCell cell = row.getCell(i);
                cell.setCellType(CellType.STRING);
                headNames.add(cell.getStringCellValue());
            }
            // head2Attr 的表头信息 key 值，要包含于 file 的表头
            for (String key : keys) {
                // 提供的表头名称不在 excel 表格中
                if (!headNames.contains(key)) {
                    return null;
                }
            }
            int index = 0;
            // 信息匹配后，需要对 head2Attr 进行重排，例如： 1  attrName
            // 存储有序的表格列和表头名称已经属性名的对应关系，因为需要的下标是表格的下标，所以使用 headNames 进行遍历
            for (String name : headNames) {
                if (keys.contains(name)) {
                    String methodName = head2Attr.get(name);
                    column2Method.put(index, methodName);
                }
                index++;
            }


            // 需要获取的总列数
            int lastNum = keys.size();
            int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
            for (int i = 1; i < physicalNumberOfRows; i++) {
                // 获取数据行的数据
                row = sheet.getRow(i);
                // 创建实体对象
                T entity = basic.newInstance();
                Class clazz = entity.getClass();

                HSSFRow hssfRow = sheet.getRow(i);
                // 创建的文件会在已有信息最后出现很多整行空白的内容，用record 记录当前行空白的列数
                // 当整行的内容都为空白表示文件结束
                int record = 0;
                int lastCellNum = hssfRow.getLastCellNum();
                Set<Integer> effectiveIndexs = column2Method.keySet();
                for (Integer effectiveIndex : effectiveIndexs) {
                    HSSFCell cell = row.getCell(effectiveIndex);
                    if (cell == null || cell.getCellType() == CellType.BLANK) {
                        record++;
                        continue;
                    }
                    cell.setCellType(CellType.STRING);
                    String cellValue = cell.getStringCellValue();
                    // 获取对应的 set 方法的参数类型，然后将 cellValue 转换成对应的类型，再调用对应的 set 方法设置属性值
                    String methodName = column2Method.get(effectiveIndex);
                    // 获取对应的 get 方法
                    Method method = clazz.getMethod("get" + methodName);
                    // 获取第一个参数的数据类型
                    String returnType = method.getReturnType().getSimpleName();
                    // 转换数据类型并执行 set 方法
                    switch (returnType) {
                        case "String":
                            method = clazz.getMethod("set" + methodName, String.class);
                            method.invoke(entity, cellValue);
                            break;
                        case "byte":
                        case "Byte":
                            method = clazz.getMethod("set" + methodName, Byte.class);
                            method.invoke(entity, Byte.parseByte(cellValue));
                            break;
                        case "int":
                        case "Integer":
                            method = clazz.getMethod("set" + methodName, Integer.class);
                            method.invoke(entity, Integer.parseInt(cellValue));
                            break;
                        case "long":
                        case "Long":
                            method.invoke(entity, Long.parseLong(cellValue));
                            break;
                        case "float":
                        case "Float":
                            method = clazz.getMethod("set" + methodName, Long.class);
                            method.invoke(entity, Float.parseFloat(cellValue));
                            break;
                        case "double":
                        case "Double":
                            method = clazz.getMethod("set" + methodName, Double.class);
                            method.invoke(entity, Double.parseDouble(cellValue));
                            break;
                        case "Date":
                            method = clazz.getMethod("set" + methodName, Date.class);
                            method.invoke(entity, dateFormat.parse(cellValue));
                            break;
                    }
                }
                if (record != lastCellNum) {
                    entities.add(entity);
                } else {
                    break;
                }
            }
            return entities;
        } catch (IOException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (InstantiationException e) {
            e.printStackTrace();
        } catch (NoSuchMethodException e) {
            e.printStackTrace();
        } catch (InvocationTargetException e) {
            e.printStackTrace();
        } catch (ParseException e) {
            e.printStackTrace();
        }

        return null;
    }
```

