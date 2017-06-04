package com.shushanfx.poi;

import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFTextParagraph;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.PAPBinTable;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.sl.usermodel.Shape;
import org.apache.poi.sl.usermodel.Slide;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.util.List;

/**
 * Created by dengjianxin on 2017/6/2.
 */
public final class ChineseUtils {
    /**
     * 处理中文
     * @param slide
     */
    public static void forChinese(XSLFSlide slide){
        XSLFTextShape[] shapes = slide.getPlaceholders();
        for (int j = 0; j < shapes.length; j++) {
            XSLFTextShape shape = shapes[j];
            if (shape.getTextParagraphs() != null) {
                shape.getTextParagraphs().forEach(item -> {
                    if (item.getTextRuns() != null) {
                        item.getTextRuns().forEach(run -> {
                            run.setFontFamily("宋体");
                        });
                    }
                });
            }
        }
    }

    /**
     * 处理中文
     * @param slide
     */
    public static void forChinese(HSLFSlide slide){
        List<List<HSLFTextParagraph>> list = slide.getTextParagraphs();
        list.forEach(array -> {
            array.forEach(item -> {
                if(item !=null && item.getTextRuns() !=null){
                    item.getTextRuns().forEach(run -> {
                        run.setFontFamily("宋体");
                    });
                }
            });
        });
    }

    public static void forChinese(XWPFDocument document){
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for(XWPFParagraph paragraph : paragraphs){
            List<XWPFRun> runs = paragraph.getRuns();
            if(runs!=null && runs.size() > 0){
                runs.forEach(run -> {
                    run.setFontFamily("宋体");
                });
            }
        }
    }
}
