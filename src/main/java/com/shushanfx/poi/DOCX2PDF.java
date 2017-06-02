package com.shushanfx.poi;

import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;

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
            PdfOptions options = PdfOptions.create();
            PdfConverter.getInstance().convert(document, out, options);
        } finally {
            input.close();
            out.close();
        }
    }
}
