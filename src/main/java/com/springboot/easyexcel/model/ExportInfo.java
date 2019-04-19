package com.springboot.easyexcel.model;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;
import lombok.Data;

/**
 * 导出 Excel 时使用的映射实体类，Excel 模型
 * @auther liupeiqing
 * @date 2019/4/19 13:25
 */
@Data
public class ExportInfo extends BaseRowModel {

    @ExcelProperty(value = "name",index = 0)
    private String name;

    @ExcelProperty(value = "年龄",index = 1)
    private String age;

    @ExcelProperty(value = "邮箱",index = 2)
    private String email;

    @ExcelProperty(value = "地址",index = 3)
    private String address;
}
