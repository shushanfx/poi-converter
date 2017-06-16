package com.shushanfx.poi;

import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFTextParagraph;
import org.apache.poi.hslf.usermodel.HSLFTextRun;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.PAPBinTable;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.sl.usermodel.Shape;
import org.apache.poi.sl.usermodel.Slide;
import org.apache.poi.sl.usermodel.TextRun;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;

import java.util.List;

/**
 * Created by dengjianxin on 2017/6/2.
 */
public final class ChineseUtils {
    /**
     * 处理中文
     * @param slide XSLFSlide a page.
     */
    public static void forChinese(XSLFSlide slide){
        XSLFTextShape[] shapes = slide.getPlaceholders();
        for (int j = 0; j < shapes.length; j++) {
            XSLFTextShape shape = shapes[j];
            List<XSLFTextParagraph> textParagraphs = shape.getTextParagraphs();
            if (textParagraphs != null) {
                for (int i = 0; i < textParagraphs.size(); i++) {
                    XSLFTextParagraph item = textParagraphs.get(i);
                    List<XSLFTextRun> runs = item.getTextRuns();
                    if(runs!=null && runs.size() > 0){
                        for (int k = 0; k < runs.size(); k++) {
                            XSLFTextRun run = runs.get(k);
                            run.setFontFamily(getFont());
                        }
                    }
                }
            }
        }
    }

    /**
     * 处理中文
     * @param slide
     */
    public static void forChinese(HSLFSlide slide){
        List<List<HSLFTextParagraph>> list = slide.getTextParagraphs();
        for(List<HSLFTextParagraph> paragraphs: list){
            for(HSLFTextParagraph paragraph : paragraphs){
                if(paragraph!=null){
                    List<HSLFTextRun> runs = paragraph.getTextRuns();
                    if(runs!=null && runs.size() > 0){
                        for(HSLFTextRun run : runs){
                            run.setFontFamily(getFont());
                        }
                    }
                }
            }
        }
    }

    public static void forChinese(XWPFDocument document){
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for(XWPFParagraph paragraph : paragraphs){
            List<XWPFRun> runs = paragraph.getRuns();
            if(runs!=null && runs.size() > 0){
                for (int i = 0; i < runs.size(); i++) {
                    XWPFRun run = runs.get(i);
                    if(run!=null){
                        run.setFontFamily(getFont());
                    }
                }
            }
        }
    }

    public static String getFont() {
        if(OsUtils.isWindows()){
            return "宋体";
        }
        else if(OsUtils.isMacOS() || OsUtils.isMacOSX()){
            return "Arial";
        }
        else{
            return "MS Sans Serif";
        }
    }
}
