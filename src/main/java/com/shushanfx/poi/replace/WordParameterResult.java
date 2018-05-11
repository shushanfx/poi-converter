package com.shushanfx.poi.replace;

import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * Created by dengjianxin on 2018/5/11.
 */
public abstract class WordParameterResult implements IWordParameterResult {
    private String key = "";

    @Override
    public void setKey(String key) {
        this.key = key;
    }

    @Override
    public String getKey() {
        return key;
    }

    protected void replaceText(XWPFRun element, String newValue){
        String text = element.getText(0);
        if(text != null){
            text = text.replace(key, newValue);
            element.setText(text, 0);
        }
    }
}
