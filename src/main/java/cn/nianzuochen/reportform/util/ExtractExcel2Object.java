package cn.nianzuochen.reportform.util;


import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellType;
import org.springframework.web.multipart.MultipartFile;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.lang.reflect.ParameterizedType;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * ToDo 将 excel 表格中的信息提取到指定的对象中
 */
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
}
