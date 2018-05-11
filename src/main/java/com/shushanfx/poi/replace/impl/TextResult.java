package com.shushanfx.poi.replace.impl;

import com.shushanfx.poi.replace.IWordParameterResult;
import com.shushanfx.poi.replace.WordParameterResult;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * Created by dengjianxin on 2018/5/11.
 */
public class TextResult extends WordParameterResult{
    private Object value = null;

    public TextResult(String key, Object value){
        this.setKey(key);
        this.value = value;
    }

    @Override
    public void handle(XWPFDocument document, XWPFRun element) {
        String text = value != null ? value.toString() : "";
        this.replaceText(element, text);
    }
}
