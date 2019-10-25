package cn.nianzuochen.reportform.util;

import cn.nianzuochen.reportform.dao.Student;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * ToDo 文件样式，文件属性相关信息
 */

public class Export2Excel {
    private final HSSFWorkbook workbook = new HSSFWorkbook();
    private String fileName;
    private LinkedHashSet<String> headName;
    private final SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");

    public Export2Excel(String fileName) {
        this.fileName = fileName;
    }

    /**
     * 创建一个 excel 文件
     * @param sheetName excel 文件的第一个 sheet 名称
     * @return
     */
    public HSSFSheet createSheet (String sheetName) {
        return workbook.createSheet(sheetName);
    }

    /**
     * 创建一个新的 sheet
     * @param sheetName
     * @return
     */
    public HSSFSheet createNewSheet(String sheetName) {
        return workbook.createSheet(sheetName);
    }

    /**
     * 返回创建的 workbook，然后可以调用相关的方法来配置信息，否则使用默认配置
     * @return
     */
    public HSSFWorkbook getWorkbook() {
        return workbook;
    }

    /**
     * ToDo 添加表头的样式
     * 设置表头
     * @param sheet 目标 sheet
     * @param head 表头名称
     * @return
     */
    public void setHead(HSSFSheet sheet, LinkedHashSet<String> head) {
//        sheet = workbook.createSheet(sheet.getSheetName());
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
            cell.setCellValue(s);
//            System.out.println("i: " + i + ", head: " + s);
        }
    }

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

    /**
     * ToDo 增加一些验证
     * Key 表示表头名称，value 则是实际的数据，相当于给出了每一列的数据
     * @param sheet
     * @param infos 需要保证每个 List 中信息条数相同
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
}
