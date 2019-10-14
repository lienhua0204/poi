package com.baizhi.poi;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

@RunWith(SpringRunner.class)
@SpringBootTest
public class PoiApplicationTests {


    @Test
    public void contextLoads() {
        //创建一个工作簿
        HSSFWorkbook sheets = new HSSFWorkbook();
        //创建工作表
        HSSFSheet sheet = sheets.createSheet("测试");
        //创建第一行
        HSSFRow row = sheet.createRow(0);
        //创建第一个单元格
        HSSFCell cell = row.createCell(0);
        //给单元格填数据
        cell.setCellValue("第一个单元格");

        try {
            sheets.write(new FileOutputStream(new File("D:/a.xls")));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    @Test
    public void contextLoads1() {
        //创建一个工作簿
        HSSFWorkbook sheets = new HSSFWorkbook();

        //创建工作表
        HSSFSheet sheet = sheets.createSheet("测试");
        //先创建标题行
        HSSFRow row = sheet.createRow(0);
        String[] str = {"id", "姓名", "性别", "生日"};
        for (int i = 0; i < str.length; i++) {
            String s = str[i];
        }

        try {
            sheets.write(new FileOutputStream(new File("D:/a.xls")));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
