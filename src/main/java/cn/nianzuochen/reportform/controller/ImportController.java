package cn.nianzuochen.reportform.controller;

import cn.nianzuochen.reportform.dao.Student;
import cn.nianzuochen.reportform.service.StudentService;
import cn.nianzuochen.reportform.util.ExtractExcel2Object;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

@RestController
@RequestMapping("import")
public class ImportController {
    @Autowired
    StudentService studentService;

    @PostMapping("student")
    public void importStudent(MultipartFile file) {
        Class studentClass = new Student().getClass();
        ExtractExcel2Object<Student> extract = new ExtractExcel2Object<>(studentClass);
        List<Student> studentList = extract.extract(file);
        for (Student student : studentList) {
            studentService.insert(student);
            System.out.println(student);
        }
    }
}
