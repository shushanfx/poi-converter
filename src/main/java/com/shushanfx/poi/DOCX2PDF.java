package com.shushanfx.poi;

import com.lowagie.text.Font;
import com.lowagie.text.FontFactory;
import com.lowagie.text.pdf.BaseFont;
import fr.opensagres.xdocreport.itext.extension.font.IFontProvider;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.awt.*;
import java.io.*;
import java.net.URISyntaxException;

/**
 * Created by dengjianxin on 2017/6/2.
 */
public class DOCX2PDF implements POIConverter {
    @Override
    public void convert(String src, String dst) throws IOException {
        InputStream input = new FileInputStream(src);
        OutputStream out = new FileOutputStream(dst);
        try{
            XWPFDocument document = new XWPFDocument(input);
            ChineseUtils.forChinese(document);
            PdfOptions pdfOptions = PdfOptions.create();
            PdfConverter.getInstance().convert(document, out, pdfOptions);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            input.close();
            out.close();
        }
    }
}
