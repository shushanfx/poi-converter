package com.shushanfx.poi.replace.impl;

import com.shushanfx.poi.replace.IWordParameterResult;
import com.shushanfx.poi.replace.WordParameterResult;
import org.apache.poi.xwpf.converter.core.utils.XWPFTableUtil;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;

import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;

/**
 * Created by dengjianxin on 2018/5/11.
 */
public class TableResult extends WordParameterResult {
    private Map<String, List<Object>> tables = null;

    public TableResult(String key, Map<String, List<Object>> tables){
        this.setKey(key);
        this.tables = tables;
    }

    private int getRowCount(){
        int iMax = 0;
        if(this.tables!=null){
            for(Map.Entry<String, List<Object>> entry : this.tables.entrySet()){{
                List<Object> value = entry.getValue();
                if(value != null && value.size() > iMax){
                    iMax = value.size();
                }
            }}
        }
        return iMax;
    }

    private XWPFTableRow getOrCreate(XWPFTable table, int i){
        XWPFTableRow row = table.getRow(i);
        if(row == null){
            row = table.createRow();
        }
        return row;
    }

    private XWPFTableCell getOrCreate(XWPFTableRow row, int i){
        XWPFTableCell cell = row.getCell(i);
        if(cell == null){
            cell = row.addNewTableCell();
        }
        return cell;
    }

    @Override
    public void handle(XWPFDocument document, XWPFRun element) {
        // 替换原始字符
        replaceText(element, "");
        if(this.tables == null){
            return ;
        }
        XWPFParagraph paragraph = (XWPFParagraph) element.getParent();
        XmlCursor cursor = paragraph.getCTP().newCursor();
        XWPFTable table = document.insertNewTbl(cursor);
        XWPFTableRow fRow = getOrCreate(table, 0);
        Set<String> keys = this.tables.keySet();
        int iKey = 0;
        for(String key : keys){
            XWPFTableCell cell = fRow.getCell(iKey);
            if(cell == null){
                cell = fRow.addNewTableCell();
            }
            cell.setText(key);
            iKey ++;
        }

        int iMax = getRowCount();
        for(int i = 0; i < iMax; i++){
            XWPFTableRow row = getOrCreate(table, i + 1);
            iKey = 0;
            for(String key: keys){
                List<Object> list = this.tables.get(key);
                XWPFTableCell cell = getOrCreate(row, iKey);
                if(list != null && i < list.size() ){
                    cell.setText(list.get(i).toString());
                }
                else{
                    cell.setText("-");
                }
                iKey ++;
            }
        }
    }
}
