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



1. 创建文件
2. 设置文件样式
3. 填充数据
4. 返回数据



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