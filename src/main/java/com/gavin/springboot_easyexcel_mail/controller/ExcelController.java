package com.gavin.springboot_easyexcel_mail.controller;

import com.gavin.springboot_easyexcel_mail.service.ExcelService;
import com.gavin.springboot_easyexcel_mail.util.easyExcel.model.ExportInfo;
import com.gavin.springboot_easyexcel_mail.util.easyExcel.model.ImportInfo;
import com.gavin.springboot_easyexcel_mail.util.easyExcel.util.ExcelUtil;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;


@RestController
@RequestMapping(value = "/excel")
public class ExcelController {
    @Autowired
    private ExcelService excelService;

    @GetMapping(value = "/sendExcel")
    public String sendExcel(@RequestParam(value = "address") String address){
        excelService.sendExcel(address);
        return "发送成功";
    }



    /**
     * 读取 Excel（允许多个 sheet）
     */
    @PostMapping(value = "/readExcelWithSheets")
    public Object readExcelWithSheets(MultipartFile excel) {
        return ExcelUtil.readExcel(excel, new ImportInfo());
    }

    /**
     * 读取 Excel（指定某个 sheet）
     */
    @PostMapping(value = "/readExcel")
    public Object readExcel(MultipartFile excel, int sheetNo,
                            @RequestParam(defaultValue = "1") int headLineNum) {
        return ExcelUtil.readExcel(excel, new ImportInfo(), sheetNo, headLineNum);
    }

    /**
     * 导出 Excel（一个 sheet）
     */
    @GetMapping(value = "/writeExcel")
    public void writeExcel(HttpServletResponse response) throws IOException {
        List<ExportInfo> list = getList();
        String fileName = "一个 Excel 文件";
        String sheetName = "第一个 sheet";

        ExcelUtil.writeExcel(response, list, fileName, sheetName, new ExportInfo());

    }

    /**
     * 导出 Excel（多个 sheet）
     */
    @GetMapping(value = "/writeExcelWithSheets")
    public void writeExcelWithSheets(HttpServletResponse response) throws IOException {
        List<ExportInfo> list = getList();
        String fileName = "一个 Excel 文件";
        String sheetName1 = "第一个 sheet";
        String sheetName2 = "第二个 sheet";
        String sheetName3 = "第三个 sheet";

        ExcelUtil.writeExcelWithSheets(response, list, fileName, sheetName1, new ExportInfo())
                .write(list, sheetName2, new ExportInfo())
                .write(list, sheetName3, new ExportInfo())
                .finish();
    }

    private List<ExportInfo> getList() {
        List<ExportInfo> list = new ArrayList<>();
        ExportInfo model1 = new ExportInfo();
        model1.setName("gavin");
        model1.setAge("19");
        model1.setAddress("南京");
        model1.setEmail("xunyegege@gmail.com");
        list.add(model1);
        ExportInfo model2 = new ExportInfo();
        model2.setName("zmx");
        model2.setAge("20");
        model2.setAddress("南京");
        model2.setEmail("gavincoder@163.com");
        list.add(model2);
        return list;
    }


}
