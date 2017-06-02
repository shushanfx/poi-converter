package com.shushanfx.poi;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Font;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;

import java.io.*;

/**
 * Created by dengjianxin on 2017/6/2.
 */
public class DOC2PDF extends DOCX2PDF {
    @Override
    public void convert(String src, String dst) throws IOException {
        try {
            super.convert(src, dst);
            return;
        } catch (Exception e) {
            System.out.println("Can't convert with XWPF.");
        }
        oldConvert(src, dst);


    }

    public void newConvert(String src, String dst) throws IOException {

    }

    public void oldConvert(String src, String dst) throws IOException {
        HWPFDocument doc = null;
        WordExtractor we = null;
        Document pdfDocument = null;
        PdfWriter pdfWriter = null;
        try {
            doc = new HWPFDocument(new FileInputStream(src));
            we = new WordExtractor(doc);
            pdfDocument = new Document();
            pdfWriter = PdfWriter.getInstance(pdfDocument, new FileOutputStream(dst));

            pdfWriter.open();
            pdfDocument.open();
            for (String text : we.getParagraphText()) {
                BaseFont bf = BaseFont.createFont("STSong-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
                Font font = new Font(bf, 14);
                pdfDocument.add(new Paragraph(text, font));
            }
        } catch (DocumentException e) {
            e.printStackTrace();
        } finally {
            if (doc != null) {
                doc.close();
            }
            if (we != null) {
                we.close();
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
