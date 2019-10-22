package cn.nianzuochen.reportform;

import org.mybatis.spring.annotation.MapperScan;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import springfox.documentation.swagger2.annotations.EnableSwagger2;

@SpringBootApplication
@EnableSwagger2
@MapperScan("cn.nianzuochen.reportform.mapper")
public class ReportFormApplication {
    public static void main(String[] args) {
        SpringApplication.run(ReportFormApplication.class, args);
    }
}
