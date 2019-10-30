package com.gavin.springboot_easyexcel_mail.util.easyExcel.model;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;
import lombok.Data;

/**

 * @Description 导入 Excel 时使用的映射实体类，Excel 模型

 */
@Data
public class ImportInfo extends BaseRowModel {
    @ExcelProperty(index = 0)
    private String name;

    @ExcelProperty(index = 1)
    private String age;

    @ExcelProperty(index = 2)
    private String email;


}
