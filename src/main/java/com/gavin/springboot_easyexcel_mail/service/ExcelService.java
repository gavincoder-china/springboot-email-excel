package com.gavin.springboot_easyexcel_mail.service;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import com.gavin.springboot_easyexcel_mail.util.easyExcel.model.ExportInfo;
import com.gavin.springboot_easyexcel_mail.util.mail.MailUtil;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.scheduling.annotation.Async;
import org.springframework.stereotype.Service;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

/**
 * **********************************************************
 *
 * @Project: 异步发送邮件
 * @Author : Gavincoder
 * @Mail : xunyegege@gmail.com
 * @Github : https://github.com/xunyegege
 * @ver : version 1.0
 * @Date : 2019-10-30 16:17
 * @description:
 ************************************************************/
@Service
@Async
public class ExcelService {


    @Autowired
    private MailUtil mailUtil;

    /**
     * @return void
     * @throws
     * @description 获取的参数为收件人的地址
     * 步骤:  异步操作,无需返回值
     * 1:从数据库拿出值
     * 2:将值赋给model
     * 3:将model打包成list
     * 4:生成excel到本地
     * 5:将该excel的地址给邮件方法
     * 6:成功
     * @author Gavin
     * @date 2019-10-30 16:19
     * @since
     */
    @Async
    public void sendExcel(String emailAddress) {

        //生成excel
        try {
            String filePath = createExcel();
            //发送邮件
            mailUtil.sendAttachmentsMail(emailAddress, "gavin", "easyExcel测试", filePath);
        }
        catch (Exception e) {
            e.printStackTrace();
        }


    }

    /**
     * @return 返回文件的路径+名称
     * @throws
     * @description 创建Excel ,返回文件路径
     * @author Gavin
     * @date 2019-10-30 19:24
     * @since
     */
    private String createExcel() throws Exception {


        // 文件输出位置+名称(时间格式)
        DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss:SSS");
        String format = dateTimeFormatter.format(LocalDateTime.now());

        //这个路径改成你自己的
        String filePath = "/Users/gavincoder/Desktop/" + format + ".xlsx";

        OutputStream out = new FileOutputStream(filePath);

        ExcelWriter writer = EasyExcelFactory.getWriter(out);

        // 写仅有一个 Sheet 的 Excel 文件, 此场景较为通用
        Sheet sheet1 = new Sheet(1, 0, ExportInfo.class);

        // 第一个 sheet 名称
        sheet1.setSheetName("FirstSheet");

        // 写数据到 Writer 上下文中
        // 入参1: 创建要写入的模型数据
        // 入参2: 要写入的目标 sheet

        //  这边你也可以从数据库拿数据,然后foreach赋值给model对象
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
        model2.setAddress("南京杭州");
        model2.setEmail("gavincoder@163.com");
        list.add(model2);

        writer.write(list, sheet1);

        // 将上下文中的最终 outputStream 写入到指定文件中
        writer.finish();

        // 关闭流
        out.close();
        return filePath;
    }
}
