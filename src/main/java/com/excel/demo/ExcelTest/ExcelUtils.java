package com.excel.demo.ExcelTest;

import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.OutputStream;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author guochunyuan
 * @create on  2018-06-07 14:04
 */
public class ExcelUtils {
    public static void expoortExcelx(String title, String[] headers, String[] columns,
                              List<String> list, OutputStream out, String pattern) throws NoSuchMethodException, Exception{
        //创建工作薄
        XSSFWorkbook workbook=new XSSFWorkbook();
        //创建表格
        Sheet sheet=workbook.createSheet(title);
        //设置默认宽度
        sheet.setDefaultColumnWidth(25);
        //创建样式
        XSSFCellStyle style=workbook.createCellStyle();
        //设置样式
        style.setFillForegroundColor(IndexedColors.GOLD.index);
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setBorderRight(CellStyle.BORDER_THIN);
        style.setBorderTop(CellStyle.BORDER_THIN);
        //生成字体
        XSSFFont font=workbook.createFont();
        font.setColor(IndexedColors.VIOLET.index);
        font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
        //应用字体
        style.setFont(font);

        //自动换行
        style.setWrapText(true);
        //声明一个画图的顶级管理器
        Drawing drawing=(XSSFDrawing) sheet.createDrawingPatriarch();
        //表头的样式
        XSSFCellStyle titleStyle=workbook.createCellStyle();//样式对象
        titleStyle.setAlignment(CellStyle.ALIGN_CENTER_SELECTION);//水平居中
        titleStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        //设置字体
        XSSFFont titleFont=workbook.createFont();
        titleFont.setFontHeightInPoints((short)15);
        titleFont.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);//粗体
        titleStyle.setFont(titleFont);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, headers.length-1));
        //指定合并区域
        Row rowHeader = sheet.createRow(0);
        //XSSFRow rowHeader=sheet.createRow(0);
        Cell cellHeader=rowHeader.createCell(0);
        XSSFRichTextString textHeader=new XSSFRichTextString(title);
        cellHeader.setCellStyle(titleStyle);
        cellHeader.setCellValue(textHeader);

        Row row=sheet.createRow(1);
        for(int i=0;i<headers.length;i++){
            Cell cell=row.createCell(i);
            cell.setCellStyle(style);
            XSSFRichTextString text=new XSSFRichTextString(headers[i]);
            cell.setCellValue(text);
        }
        //遍历集合数据，产生数据行
        if(list!=null&&list.size()>0){
            int index=2;
            for(String t:list){
                row=sheet.createRow(index);
                index++;
                for(short i=0;i<columns.length;i++){
                    Cell cell=row.createCell(i);
                    String filedName=columns[i];
                    String getMethodName="get"+filedName.substring(0,1).toUpperCase()
                            +filedName.substring(1);
                    Class tCls=t.getClass();
                    Method getMethod=tCls.getMethod(getMethodName,new Class[]{});
                    Object value=getMethod.invoke(t, new Class[]{});
                    String textValue=null;
                    if(value==null){
                        textValue="";
                    }else if(value instanceof Date){
                        Date date=(Date)value;
                        SimpleDateFormat sdf = new SimpleDateFormat(pattern);
                        textValue = sdf.format(date);
                    }else if(value instanceof byte[]){
                        row.setHeightInPoints(80);
                        sheet.setColumnWidth(i, 35*100);
                        byte[] bsValue=(byte[])value;
                        XSSFClientAnchor anchor=new XSSFClientAnchor(0,0,1023,255,6,index,6,index);
                        anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
                        drawing.createPicture(anchor, workbook.addPicture(bsValue, XSSFWorkbook.PICTURE_TYPE_JPEG));
                    }else{
                        // 其它数据类型都当作字符串简单处理
                        textValue=value.toString();
                    }

                    if(textValue!=null){
                        Pattern p = Pattern.compile("^//d+(//.//d+)?$");
                        Matcher matcher = p.matcher(textValue);
                        if (matcher.matches()) {
                            // 是数字当作double处理
                            cell.setCellValue(Double.parseDouble(textValue));
                        } else {
                            XSSFRichTextString richString = new XSSFRichTextString(
                                    textValue);
                            // HSSFFont font3 = workbook.createFont();
                            // font3.setColor(HSSFColor.BLUE.index);
                            // richString.applyFont(font3);
                            cell.setCellValue(richString);
                        }
                    }
                }
            }
        }
        workbook.write(out);
        workbook.close();
        out.close();
    }
}
