package com.shushanfx.poi;

import java.io.File;
import java.io.IOException;

/**
 * Created by dengjianxin on 2017/6/2.
 */
public class POIMain {
    public static void main(String[] args) throws IOException {
        String src = args[0];
        String dst = args[1];
        if (isBlank(src) || isBlank(dst)) {
            usage();
            return;
        }
        convert(src, dst);
    }

    public static void convert(String src, String dst) throws IOException {
        long start = System.currentTimeMillis();

        int index = src.lastIndexOf(".");
        String ext = null;
        if (index >= 0 && index < src.length() - 1) {
            ext = src.substring(index + 1);
        }

        POIConverter converter = null;
        if ("docx".equalsIgnoreCase(ext)) {
            converter = new DOCX2PDF();
        }
        else if("doc".equalsIgnoreCase(ext)){
            converter = new DOC2PDF();
        }
        else if("pptx".equalsIgnoreCase(ext)){
            converter = new PPTX2PDF();
        }
        else if("ppt".equalsIgnoreCase(ext)){
            converter = new PPT2PDF();
        }
        if (converter != null) {
            guaranteeParent(dst);
            converter.convert(src, dst);
        }
        else{
            usage();
        }

        System.out.println("Convert from " + src + " to " + dst + " in " + (System.currentTimeMillis() - start) + " ms");
    }

    public static boolean isBlank(String str) {
        return str == null || "".equals(str.trim());
    }

    public static void usage() {
        System.out.println("Usage: POIMain fromFile toFile\n Only .doc, .docx, .ppt, .pptx are supported.");
    }

    private static void guaranteeParent(String path){
        File file = new File(path);
        File parentFile = file.getParentFile();
        if(parentFile!=null){
            if(!parentFile.exists()){
                parentFile.mkdirs();
            }
        }
    }
}
