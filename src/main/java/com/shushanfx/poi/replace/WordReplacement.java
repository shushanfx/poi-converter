package com.shushanfx.poi.replace;

import com.shushanfx.poi.replace.impl.TableResult;
import com.shushanfx.poi.replace.impl.TextResult;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import java.io.*;
import java.util.*;

/**
 * Created by dengjianxin on 2018/5/11.
 */
public class WordReplacement{
    private List<IWordParameterResult> replace = null;
    private XWPFDocument document = null;

    public WordReplacement(XWPFDocument document){
        this.replace = new ArrayList<IWordParameterResult>();
        this.document = document;
    }

    private WordReplacement(){}

    public void add(IWordParameterResult item){
        replace.add(item);
    }

    public void handle(){
        if(this.document != null){
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            for(XWPFParagraph paragraph : paragraphs){
                List<XWPFRun> runs = paragraph.getRuns();
                for(XWPFRun run : runs){
                    doHandle(run);
                }
            }

            List<XWPFTable> tables = document.getTables();
            for(XWPFTable table : tables){

            }
        }
    }

    private void doHandle(XWPFRun run){
        String text = run.getText(0);
        if(text != null){
            for(IWordParameterResult item : replace){
                if(item instanceof WordParameterResult){
                    WordParameterResult result = (WordParameterResult) item;
                    String key = result.getKey();
                    if(text.indexOf(key) != -1){
                        result.handle(run.getDocument(), run);
                    }
                }
            }
        }
    }


    public XWPFDocument getDocument() {
        return document;
    }


    public static void main(String[] args) throws IOException {
        XWPFDocument document = new XWPFDocument(new FileInputStream(new File("res/replace.docx")));
        WordReplacement replacement = new WordReplacement(document);
        replacement.add(new TextResult("${合同编号}", "XXXX"));
        replacement.add(new TextResult("${租客名称}", "shushanfx"));
        replacement.add(new TextResult("${通讯地址}", "北京市海淀区北京市" +
                "海淀区北京市北京市海淀区北京市海淀区" +
                "北京市北京市海淀区北京市海淀区北京市" +
                "北京市海淀区北京市海淀区北京市北京市" +
                "海淀区北京市海淀区北京市北京市海淀区" +
                "北京市海淀区北京市"));
        replacement.add(new TextResult("${租客联系人电话}", "18511453850"));

        Map<String, List<Object>> map = new HashMap<String, List<Object>>();
        List<Object> list = new ArrayList<>();
        list.add(1);
        list.add(new Date());
        list.add(new String("Hello"));
        map.put("序号", list);
        map.put("姓名", list);
        map.put("描述", list);
        map.put("测试", list);
        replacement.add(new TableResult("${table}", map));

        replacement.handle();


        document.write(new FileOutputStream("dist/replace.docx"));
    }
}

