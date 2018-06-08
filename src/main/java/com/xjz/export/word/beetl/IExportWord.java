package com.xjz.export.word.beetl;

import java.util.Map;

/**
 * @author: Gibbons
 * @create: 2018/06/08 19:04
 * @description:
 **/
public interface IExportWord {
    void exportWord(String outPath, Map<String, Object> dataMap) throws Exception;
}
