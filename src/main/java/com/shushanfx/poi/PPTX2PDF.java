package com.shushanfx.poi;

import com.itextpdf.text.*;
import com.itextpdf.text.Image;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import java.awt.*;
import java.awt.geom.AffineTransform;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.*;

/**
 * Created by dengjianxin on 2017/6/2.
 */
public class PPTX2PDF implements POIConverter {
    @Override
    public void convert(String src, String dst) throws IOException {
        InputStream input = new FileInputStream(src);
        OutputStream output = new FileOutputStream(dst);
        PdfWriter pdfWriter = null;
        Document pdfDocument = null;
        XMLSlideShow ppt = null;

        try {
            ppt = new XMLSlideShow(input);
            double zoom = 2;
            AffineTransform at = new AffineTransform();
            at.setToScale(zoom, zoom);
            pdfDocument = new Document();
            pdfWriter = PdfWriter.getInstance(pdfDocument, output);

            PdfPTable table = new PdfPTable(1);
            pdfWriter.open();
            pdfDocument.open();

            Dimension pageSize = ppt.getPageSize();
            java.util.List<XSLFSlide> slides = ppt.getSlides();
            pdfDocument.setPageSize(new Rectangle((int) pageSize.getWidth(), (int) pageSize.getHeight()));
            for (int i = 0; i < slides.size(); i++) {
                XSLFSlide slide = slides.get(i);
                ChineseUtils.forChinese(slide);
                BufferedImage img = new BufferedImage((int) Math.ceil(pageSize.width * zoom), (int) Math.ceil(pageSize.height * zoom), BufferedImage.TYPE_INT_RGB);
                Graphics2D graphics = img.createGraphics();
                graphics.setTransform(at);
                graphics.setPaint(Color.white);
                graphics.fill(new Rectangle2D.Float(0, 0, pageSize.width, pageSize.height));
                slide.draw(graphics);
                Image slideImage = com.itextpdf.text.Image.getInstance(img, null);
                table.addCell(new PdfPCell(slideImage, true));
            }
            pdfDocument.add(table);
        } catch (BadElementException e) {
            e.printStackTrace();
        } catch (DocumentException e) {
            e.printStackTrace();
        } finally {
            if (ppt != null) {
                ppt.close();
            }
            if (pdfDocument != null) {
                pdfDocument.close();
            }
            if (pdfWriter != null) {
                pdfWriter.close();
            }
        }

    }
}
