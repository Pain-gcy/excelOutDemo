package com.excel.demo.ExcelTest;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.FileOutputStream;

/**
 * @author guochunyuan
 * @create on  2018-06-07 15:25
 */
public class TestDemo {
    public static String outputFile = "D:\\test1.xls";
    public static String fileToBeRead = "D:\\test1.xls";
    public static void main(String[] args) {
        //新建excel
        try {
            // 创建新的Excel 工作簿
            HSSFWorkbook workbook = new HSSFWorkbook();

            // 在Excel工作簿中建一工作表，其名为缺省值
            // 如要新建一名为"效益指标"的工作表，其语句为：
            // HSSFSheet sheet = workbook.createSheet("效益指标");
            HSSFSheet sheet = workbook.createSheet("效益指标");
            // 在索引0的位置创建行（最顶端的行）
            HSSFRow row = sheet.createRow((short) 0);

            HSSFCell empCodeCell = row.createCell((short) 0);
            empCodeCell.setCellType(HSSFCell.CELL_TYPE_STRING);
            empCodeCell.setCellValue("员工代码");

            HSSFCell empNameCell = row.createCell((short) 1);
            empNameCell.setCellType(HSSFCell.CELL_TYPE_STRING);
            empNameCell.setCellValue("姓名");

            HSSFCell sexCell = row.createCell((short) 2);
            sexCell.setCellType(HSSFCell.CELL_TYPE_STRING);
            sexCell.setCellValue("性别");

            HSSFCell birthdayCell = row.createCell((short) 3);
            birthdayCell.setCellType(HSSFCell.CELL_TYPE_STRING);
            birthdayCell.setCellValue("出生日期");

            HSSFCell orgCodeCell = row.createCell((short) 4);
            orgCodeCell.setCellType(HSSFCell.CELL_TYPE_STRING);
            orgCodeCell.setCellValue("机构代码");

            HSSFCell orgNameCell = row.createCell((short) 5);
            orgNameCell.setCellType(HSSFCell.CELL_TYPE_STRING);
            orgNameCell.setCellValue("机构名称");

            HSSFCell contactTelCell = row.createCell((short) 6);
            contactTelCell.setCellType(HSSFCell.CELL_TYPE_STRING);
            contactTelCell.setCellValue("联系电话");

            HSSFCell zjmCell = row.createCell((short) 7);
            zjmCell.setCellType(HSSFCell.CELL_TYPE_STRING);
            zjmCell.setCellValue("助记码");
            for (int i = 1; i <= 10; i++) {
                row = sheet.createRow((short) i);
                empCodeCell = row.createCell((short) 0);
                empCodeCell.setCellType(HSSFCell.CELL_TYPE_STRING);
                empCodeCell.setCellValue("001_" + i);

                empNameCell = row.createCell((short) 1);
                empNameCell.setCellType(HSSFCell.CELL_TYPE_STRING);
                empNameCell.setCellValue("张三_" + i);

                sexCell = row.createCell((short) 2);
                sexCell.setCellType(HSSFCell.CELL_TYPE_STRING);
                sexCell.setCellValue("性别_" + i);

                birthdayCell = row.createCell((short) 3);
                birthdayCell.setCellType(HSSFCell.CELL_TYPE_STRING);
                birthdayCell.setCellValue("出生日期_" + i);

                orgCodeCell = row.createCell((short) 4);
                orgCodeCell.setCellType(HSSFCell.CELL_TYPE_STRING);
                orgCodeCell.setCellValue("机构代码_" + i);

                orgNameCell = row.createCell((short) 5);
                orgNameCell.setCellType(HSSFCell.CELL_TYPE_STRING);
                orgNameCell.setCellValue("机构名称_" + i);

                contactTelCell = row.createCell((short) 6);
                contactTelCell.setCellType(HSSFCell.CELL_TYPE_STRING);
                contactTelCell.setCellValue("联系电话_" + i);

                zjmCell = row.createCell((short) 7);
                zjmCell.setCellType(HSSFCell.CELL_TYPE_STRING);
                zjmCell.setCellValue("助记码_" + i);

            }
            // 新建一输出文件流
            FileOutputStream fOut = new FileOutputStream(outputFile);
            // 把相应的Excel 工作簿存盘
            workbook.write(fOut);
            fOut.flush();
            // 操作结束，关闭文件
            fOut.close();
            workbook.close();
            System.out.println("文件生成...");

        } catch (Exception e) {
            System.out.println("已运行 xlCreate() : " + e);
        }

        //更改数据
        try {
            FileInputStream fs = new FileInputStream("d:\\test1.xls"); // 获取d://test.xls
            POIFSFileSystem ps = new POIFSFileSystem(fs); // 使用POI提供的方法得到excel的信息
            HSSFWorkbook wb = new HSSFWorkbook(ps);
            HSSFSheet sheet = wb.getSheetAt(0); // 获取到工作表，因为一个excel可能有多个工作表
            HSSFRow row = sheet.getRow(0); // 获取第一行（excel中的行默认从0开始，所以这就是为什么，一个excel必须有字段列头），即，字段列头，便于赋值
            System.out.println(sheet.getLastRowNum() + " " + row.getLastCellNum()); // 分别得到最后一行的行号，和一条记录的最后一个单元格

            FileOutputStream out = new FileOutputStream("d:\\test1.xls"); // 向d://test.xls中写数据
            row = sheet.createRow((short) (sheet.getLastRowNum() + 1)); // 在现有行号后追加数据
            row.createCell(0).setCellValue("leilei"); // 设置第一个（从0开始）单元格的数据
            row.createCell(1).setCellValue(24); // 设置第二个（从0开始）单元格的数据

            out.flush();
            wb.write(out);
            out.close();
            // System.out.println(row.getPhysicalNumberOfCells() + "
            // " + row.getLastCellNum());
        } catch (Exception e) {
            System.out.println("已运行xlRead() : " + e);
        }


        //读取数据
        try {
            FileOutputStream fOut = new FileOutputStream("d:\\test2.xls");
            // 创建对Excel工作簿文件的引用
            HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream("d:\\test2.xls"));
            HSSFSheet sheet = workbook.getSheetAt(0);
            int i = 0;
            while (true) {
                HSSFRow row = sheet.getRow(i);
                if (row == null) {
                    break;
                }
                HSSFCell cell0 = row.getCell((short) 0);
                HSSFCell cell1 = row.getCell((short) 1);
                HSSFCell cell2 = row.getCell((short) 2);
                HSSFCell cell3 = row.getCell((short) 3);
                HSSFCell cell4 = row.getCell((short) 4);
                HSSFCell cell5 = row.getCell((short) 5);
                HSSFCell cell6 = row.getCell((short) 6);

                System.out.print(cell0.getStringCellValue());
                System.out.print("," + cell1.getStringCellValue());
                System.out.print("," + cell2.getStringCellValue());
                System.out.print("," + cell3.getStringCellValue());
                System.out.print("," + cell4.getStringCellValue());
                System.out.print("," + cell5.getStringCellValue());
                System.out.println("," + cell6.getStringCellValue());
                i++;
                row = sheet.createRow((short) i);
                HSSFCell empCodeCell = row.createCell((short) 0);
                empCodeCell = row.createCell((short) 0);
                empCodeCell.setCellType(HSSFCell.CELL_TYPE_STRING);
                empCodeCell.setCellValue("001_" + i + 9);
                workbook.write(fOut);
                fOut.flush();
                // 操作结束，关闭文件
                fOut.close();

            }
        } catch (Exception e) {
            System.out.println("已运行xlRead() : " + e);
        }
        System.exit(0);

    }
}
