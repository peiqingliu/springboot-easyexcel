package com.springboot.easyexcel.model;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;

/**
 *  导入 Excel 时使用的映射实体类，Excel 模型
 * @auther liupeiqing
 * @date 2019/4/19 13:29
 */
public class ImportInfo extends BaseRowModel {

    @ExcelProperty(index = 0)
    private String name;


    @ExcelProperty(index = 1)
    private String age;

    @ExcelProperty(index = 2)
    private String email;

    /*
    作为 excel 的模型映射，需要 setter 方法
    */
    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getAge() {
        return age;
    }

    public void setAge(String age) {
        this.age = age;
    }

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    @Override
    public String toString() {
        return "ImportInfo{" +
                "name='" + name + '\'' +
                ", age='" + age + '\'' +
                ", email='" + email + '\'' +
                '}';
    }
}
