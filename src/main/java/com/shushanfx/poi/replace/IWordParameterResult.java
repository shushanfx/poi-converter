package com.shushanfx.poi.replace;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * Created by dengjianxin on 2018/5/11.
 */
public interface IWordParameterResult {
    void setKey(String key);
    String getKey();
    void handle(XWPFDocument document, XWPFRun element);
}
