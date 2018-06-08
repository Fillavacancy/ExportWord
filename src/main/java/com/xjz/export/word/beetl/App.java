package com.xjz.export.word.beetl;

import com.xjz.export.word.beetl.impl.BTLExportWord;
import com.xjz.export.word.beetl.impl.FTLExportWord;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @author: Gibbons
 * @create: 2018/06/08 19:03
 * @description:
 **/
public class App {
    private final static int flag = 1;

    public static void main(String[] args) {
        Map<String, Object> data = getData();
        IExportWord exportWord;
        if (1 == flag) {
            exportWord = new BTLExportWord();
        } else {
            exportWord = new FTLExportWord();
        }
        try {
            exportWord.exportWord("C:\\Users\\Gibbons\\Desktop\\合同模板.doc", data);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static Map<String, Object> getData() {
        Map<String, Object> data = new HashMap<String, Object>();
        data.put("lx", "维修");
        data.put("xf", "四川观想科技股份有限公司");
        data.put("htbh", "HT20180606");
        data.put("gs", "XXXX公司");
        DateFormat bf = new SimpleDateFormat("yyyy-MM-dd");//多态
        data.put("qdsj", bf.format(new Date()));
        data.put("qddd", "成都");
        data.put("rmbdx", "壹佰元整");
        data.put("rmbxx", "100.00");
        data.put("cgsh", "采购");
        data.put("jhbsh", "计划部");
        data.put("cwsh", "财务");
        data.put("zjlsh", "总经理");
        data.put("bz", "很快地恢复都十分好看的说法");
        List<Map<String, Object>> items = new ArrayList<Map<String, Object>>();
        int i = 1;
        Map<String, Object> item1 = new HashMap<String, Object>();
        item1.put("xh", i++);
        item1.put("cpmc", "产品名称");
        item1.put("ggxh", "规格型号");
        item1.put("sl", 1);
        item1.put("dw", 100);
        item1.put("zje", 100);
        item1.put("jhrq", bf.format(new Date()));
        item1.put("zxbz", "好");
        Map<String, Object> item2 = new HashMap<String, Object>();
        item2.put("xh", i++);
        item2.put("cpmc", "产品名称");
        item2.put("ggxh", "规格型号");
        item2.put("sl", 2);
        item2.put("dw", 100);
        item2.put("zje", 200);
        item2.put("jhrq", bf.format(new Date()));
        item2.put("zxbz", "良好");
        items.add(item1);
        items.add(item2);
        data.put("items", items);
        return data;
    }
}
