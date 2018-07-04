package com.excel.demo.otherExcelOut;

import com.excel.demo.otherExcelOut.bean.Book;
import com.excel.demo.otherExcelOut.bean.Student;

import javax.swing.*;
import java.io.*;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @author guochunyuan
 * @create on  2018-07-04 9:33
 */
public class ExcelMain {
    public static void main(String[] args) {
        // 测试学生
        ExportExcel<Student> ex = new ExportExcel<Student>();
        String[] headers =
                {"学号", "姓名", "年龄", "性别", "出生日期"};
        List<Student> dataset = new ArrayList<Student>();
        dataset.add(new Student(10000001, "张三", 20, true, new Date()));
        dataset.add(new Student(20000002, "李四", 24, false, new Date()));
        dataset.add(new Student(30000003, "王五", 22, true, new Date()));
        // 测试图书
        ExportExcel<Book> ex2 = new ExportExcel<Book>();
        String[] headers2 = {"图书编号", "图书名称", "图书作者", "图书价格", "图书ISBN", "图书出版社", "封面图片"};
        List<Book> dataset2 = new ArrayList<Book>();
        try {
            BufferedInputStream bis = new BufferedInputStream(new FileInputStream("D://app/book.jpg"));
            byte[] buf = new byte[bis.available()];
            while ((bis.read(buf)) != -1) {
                //
            }
            dataset2.add(new Book(1, "jsp", "leno", 300.33f, "1234567",
                    "清华出版社", buf));
            dataset2.add(new Book(2, "java编程思想", "brucl", 300.33f, "1234567",
                    "阳光出版社", buf));
            dataset2.add(new Book(3, "DOM艺术", "lenotang", 300.33f, "1234567",
                    "清华出版社", buf));
            dataset2.add(new Book(4, "c++经典", "leno", 400.33f, "1234567",
                    "清华出版社", buf));
            dataset2.add(new Book(5, "c#入门", "leno", 300.33f, "1234567",
                    "汤春秀出版社", buf));

            OutputStream out = new FileOutputStream("D://app/student.xls");
            OutputStream out2 = new FileOutputStream("D://app/book.xls");
            ex.exportExcel( headers, dataset, out);
            ex2.exportExcel(headers2, dataset2, out2);
            JOptionPane.showMessageDialog(null, "导出成功!");
            System.out.println("excel导出成功！");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
