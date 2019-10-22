package cn.nianzuochen.reportform.service;

import cn.nianzuochen.reportform.dao.Student;
import org.springframework.http.ResponseEntity;

public interface StudentService {
    int insert(Student record);
    Student selectByPrimaryKey(Integer id);
    ResponseEntity<byte[]> exportStudentById(Integer id);
}
