package cn.nianzuochen.reportform.controller;

import cn.nianzuochen.reportform.service.StudentService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("export")
public class ExportController {
    @Autowired
    StudentService studentService;

    @ApiOperation(value="测试",produces="application/octet-stream")
    @GetMapping("loadstudentbyid/{id}")
    public ResponseEntity<byte[]> exportStudent(@PathVariable("id") Integer id) {
        return studentService.exportStudentById(id);
    }
}
