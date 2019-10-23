import cn.nianzuochen.reportform.ReportFormApplication;
import cn.nianzuochen.reportform.dao.Student;
import cn.nianzuochen.reportform.service.impl.StudentServiceImpl;
import lombok.extern.slf4j.Slf4j;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.util.Date;

@Slf4j
@RunWith(SpringRunner.class)
@SpringBootTest(classes = ReportFormApplication.class)
public class ReportFormApplicationTest {
    @Autowired
    StudentServiceImpl studentService;

    @Test
    public void testAddStudent() {
        Student student = new Student();
        student.setId(1);
        student.setName("Nianzuochen");
        student.setGender((byte)1);
        student.setBir(new Date());
        student.setTel("18804525726");
        student.setEmail("123@qq.com");
        student.setAddress("江苏省徐州市");
        student.setMajor("软件工程");
        student.setSchool("齐大");

        studentService.insert(student);
    }
}
