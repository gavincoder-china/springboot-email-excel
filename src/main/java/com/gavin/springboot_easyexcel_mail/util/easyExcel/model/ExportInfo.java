package com.gavin.springboot_easyexcel_mail.util.easyExcel.model;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;
import lombok.Data;

/**
 * @Description 导出 Excel 时使用的映射实体类，Excel 模型
 * 每行数据是一个java模型有表头----表头层级为多层级
 * https://blog.csdn.net/Pruett/article/details/87455949
 */

@Data
public class ExportInfo extends BaseRowModel {
    @ExcelProperty(value = {"姓名","姓名","姓名"} ,index = 0)
    private String name;

    @ExcelProperty(value = {"统一测试","测试","年龄"},index = 1)
    private int age;

    @ExcelProperty(value = {"统一测试","邮箱","邮箱"},index = 2)
    private String email;

    @ExcelProperty(value = {"地址","地址","地址"},index = 3)
    private String address;


}
