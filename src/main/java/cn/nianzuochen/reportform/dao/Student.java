package cn.nianzuochen.reportform.dao;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.util.Date;

@Getter
@Setter
@ToString
public class Student {
    private Integer id;

    private String name;

    private Byte gender;

    private Date bir;

    private String tel;

    private String email;

    private String address;

    private String major;

    private String school;
}