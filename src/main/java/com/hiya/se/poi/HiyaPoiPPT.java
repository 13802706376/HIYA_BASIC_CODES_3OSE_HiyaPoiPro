package com.hiya.se.poi;

import java.awt.Color;
import java.awt.Shape;
import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.UUID;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.xslf.usermodel.TextAlign;
import org.apache.poi.xslf.usermodel.VerticalAlignment;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFHyperlink;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class HiyaPoiPPT implements IPio
{
    public void doCreate(String path)
    {
        XMLSlideShow ppt = new XMLSlideShow();
        XSLFSlide slide = ppt.createSlide();// 创建幻灯片
        XSLFTextBox textBox = slide.createTextBox();
        textBox.setAnchor(new Rectangle2D.Double(10, 10, 0, 0));
        textBox.addNewTextParagraph().addNewTextRun().setText("创建幻灯片");

        Object[][] datas =
        {
                { "区域产品销售额", "", "" },
                { "区域", "总销售额(万元)", "总利润(万元)简单的表格" },
                { "江苏省", 9045, 2256 },
                { "广东省", 3000, 690 } };
        XSLFTable table = slide.createTable();// 创建表格
        table.setAnchor(new Rectangle2D.Double(10, 50, 0, 0));
        for (int i = 0; i < datas.length; i++)
        {
            XSLFTableRow tableRow = table.addRow(); // 创建表格行
            for (int j = 0; j < datas[i].length; j++)
            {
                XSLFTableCell tableCell = tableRow.addCell();// 创建表格单元格
                XSLFTextParagraph p = tableCell.addNewTextParagraph();
                XSLFTextRun tr = p.addNewTextRun();
                tr.setText(String.valueOf(datas[i][j]));

                tableCell.setFillColor(Color.getColor("0xdd7e6b"));
                p.setTextAlign(TextAlign.CENTER);
                tableCell.setVerticalAlignment(VerticalAlignment.MIDDLE);

                if (i == datas.length - 1 && j == 3 - 1)
                {
                    tr.setFontSize(16);
                    tr.setBold(true);
                    tr.setItalic(true);
                    tr.setUnderline(true);
                    tr.setFontFamily("\u5b8b\u4f53");
                    tr.setFontColor(Color.RED);
                }

                tableCell.setBorderBottom(1);
                tableCell.setBorderLeft(1);
                tableCell.setBorderTop(1);
                tableCell.setBorderRight(1);
                tableCell.setBorderBottomColor(Color.BLACK);
                tableCell.setBorderLeftColor(Color.BLACK);
                tableCell.setBorderTopColor(Color.BLACK);
                tableCell.setBorderRightColor(Color.BLACK);
            }

            tableRow.setHeight(30);
        }

        // 设置列宽
        table.setColumnWidth(0, 150);
        table.setColumnWidth(1, 150);
        table.setColumnWidth(2, 250);

        // 合并单元格
        table.mergeCells(0, 0, 0, 2);
     
        try
        {
            ppt.write(new FileOutputStream("/Users/mike/ppt8.pptx"));
            // 插入图片
            byte[] bt = FileUtils.readFileToByteArray(new File("/Users/mike/pie.png"));
            int idx = ppt.addPicture(bt, XSLFPictureData.PICTURE_TYPE_PNG);
            XSLFPictureShape pic = slide.createPicture(idx);
            pic.setAnchor(new Rectangle2D.Double(10, 200, 339, 197));

            // 创建一个文本链接
            XSLFTextBox linkText = slide.createTextBox();
            linkText.setAnchor(new Rectangle2D.Double(430, 310, 0, 0));
            XSLFTextRun r = linkText.addNewTextParagraph().addNewTextRun();
            r.setText("Apache POI");
            XSLFHyperlink link = r.createHyperlink();
            link.setAddress("http://poi.apache.org");
        } catch (IOException e)
        {
            e.printStackTrace();
        }
    }

    public void doParse(String path)
    {
        
    }
}
