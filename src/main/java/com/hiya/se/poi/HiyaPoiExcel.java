package com.hiya.se.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class HiyaPoiExcel implements IPio
{
    @Override
    public void doCreate(String path)
    {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("0");
        Row row = sheet.createRow(0);
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
        cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
        cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
        cellStyle.setBorderRight(CellStyle.BORDER_THIN);
        cellStyle.setBorderTop(CellStyle.BORDER_THIN);
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);

        row.createCell(0).setCellStyle(cellStyle);
        row.createCell(0).setCellValue("姓名");
        row.createCell(1).setCellStyle(cellStyle);
        row.createCell(1).setCellValue("年龄");
        workbook.setSheetName(0, "信息");
        FileOutputStream fileoutputStream  = null;
        try
        {
            File file = new File(path);
            fileoutputStream = new FileOutputStream(file);
            workbook.write(fileoutputStream);
        } catch (IOException e)
        {
            e.printStackTrace();
        }
        finally
        {
            FileUtils.close(fileoutputStream, "close");
            FileUtils.close(workbook, "close");
        }
    }

    @Override
    public void doParse(String path)
    {
        File file = new File(path);
        if (!file.exists())
            System.out.println("文件不存在");
        HSSFWorkbook hssfWorkbook = null;
        try
        {
            // 1.读取Excel的对象
            POIFSFileSystem poifsFileSystem = new POIFSFileSystem(new FileInputStream(file));
            // 2.Excel工作薄对象
            hssfWorkbook = new HSSFWorkbook(poifsFileSystem);
            // 3.Excel工作表对象
            HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(0);
            // 总行数
            int rowLength = hssfSheet.getLastRowNum() + 1;
            // 4.得到Excel工作表的行
            HSSFRow hssfRow = hssfSheet.getRow(0);
            // 总列数
            int colLength = hssfRow.getLastCellNum();
            // 得到Excel指定单元格中的内容
            HSSFCell hssfCell = hssfRow.getCell(0);
            // 得到单元格样式
            CellStyle cellStyle = hssfCell.getCellStyle();

            for (int i = 0; i < rowLength; i++)
            {
                // 获取Excel工作表的行
                HSSFRow hssfRow1 = hssfSheet.getRow(i);
                for (int j = 0; j < colLength; j++)
                {
                    // 获取指定单元格
                    HSSFCell hssfCell1 = hssfRow1.getCell(j);

                    // Excel数据Cell有不同的类型，当我们试图从一个数字类型的Cell读取出一个字符串时就有可能报异常：
                    // Cannot get a STRING value from a NUMERIC cell
                    // 将所有的需要读的Cell表格设置为String格式
                    if (hssfCell1 != null)
                    {
                        // hssfCell1.setCellType(CellType.STRING);
                    }

                    // 获取每一列中的值
                    System.out.print(hssfCell1.getStringCellValue() + "\t");
                }
                System.out.println();
            }
        } catch (IOException e)
        {
            e.printStackTrace();
        }
        finally
        {
            FileUtils.close(hssfWorkbook, "close");
        }
    }
}
