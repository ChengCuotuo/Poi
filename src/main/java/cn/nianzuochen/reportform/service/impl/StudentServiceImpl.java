package cn.nianzuochen.reportform.service.impl;

import cn.nianzuochen.reportform.dao.Student;
import cn.nianzuochen.reportform.mapper.StudentMapper;
import cn.nianzuochen.reportform.service.StudentService;
import cn.nianzuochen.reportform.util.Export2Excel;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;

@Service
public class StudentServiceImpl implements StudentService {
    @Autowired
    StudentMapper studentMapper;

    @Override
    public int insert(Student record) {
        return studentMapper.insert(record);
    }

    @Override
    public Student selectByPrimaryKey(Integer id) {
        return studentMapper.selectByPrimaryKey(id);
    }

    @Override
    public ResponseEntity<byte[]> exportStudentById(Integer id) {
        Student student = studentMapper.selectByPrimaryKey(id);
        Export2Excel export2Excel = new Export2Excel("Student");
        HSSFSheet sheet = export2Excel.createSheet("student");
        List<Student> list = new ArrayList<>();
        list.add(student);
        return export2Excel.addInfo(sheet, list);
    }


}
