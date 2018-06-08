package com.xjz.export.word.beetl.impl;

import com.xjz.export.word.beetl.IExportWord;
import com.xjz.export.word.beetl.constant.ConstConstant;
import freemarker.template.Configuration;
import freemarker.template.Template;

import java.io.*;
import java.util.Map;

/**
 * @author: Gibbons
 * @create: 2018/06/08 19:05
 * @description:
 **/
public class FTLExportWord implements IExportWord {

    private static Template t;

    private Configuration createConfig() throws IOException {
        Configuration configuration = new Configuration(Configuration.VERSION_2_3_23);
        configuration.setDefaultEncoding("UTF-8");
        configuration.setClassForTemplateLoading(FTLExportWord.class, ConstConstant.TEMPLATE_PACK);
        configuration.setClassicCompatible(true);
        return configuration;
    }

    private Template createTemplate() throws IOException {
        Configuration configuration = createConfig();
        Template template = configuration.getTemplate(ConstConstant.BTL_HTMB_DEMO);
        t = template;
        return t;
    }

    public void exportWord(String outPath, Map<String, Object> dataMap) throws Exception {
        if (null == t) {
            createTemplate();
        }
        File outFile = new File(outPath);
        if (!outFile.getParentFile().exists()) {
            outFile.getParentFile().mkdirs();
        }
        FileOutputStream fos = new FileOutputStream(outFile);
        OutputStreamWriter oWriter = new OutputStreamWriter(fos, "utf-8");
        Writer out = new BufferedWriter(oWriter);
        t.process(dataMap, out);
    }
}
