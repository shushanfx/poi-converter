package com.shushanfx.poi;

import java.io.IOException;

/**
 * Created by dengjianxin on 2017/6/2.
 */
public interface POIConverter {
    void convert(String src, String dst) throws IOException;
}
