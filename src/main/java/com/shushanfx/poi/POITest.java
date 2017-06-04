package com.shushanfx.poi;

import java.io.IOException;

import static com.shushanfx.poi.POIMain.*;

/**
 * Created by dengjianxin on 2017/6/2.
 */
public class POITest {
    public static void main(String[] args) throws IOException {
        convert("res/a.ppt", "dist/a.ppt.pdf");
        convert("res/a.pptx", "dist/a.pptx.pdf");
        convert("res/a.doc", "dist/a.doc.pdf");
        convert("res/a.docx", "dist/a.docx.pdf");
    }
}
