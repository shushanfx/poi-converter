package com.shushanfx.poi;

import java.io.IOException;

import static com.shushanfx.poi.POIMain.*;

/**
 * Created by dengjianxin on 2017/6/2.
 */
public class POITest {
    public static void main(String[] args) throws IOException {
        convert("a.ppt", "a.ppt.pdf");
        convert("a.pptx", "a.pptx.pdf");
        convert("res/a.doc", "dist/a.doc.pdf");
        convert("res/a.docx", "dist/a.docx.pdf");
    }
}
