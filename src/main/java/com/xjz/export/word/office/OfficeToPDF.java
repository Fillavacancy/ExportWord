package com.xjz.export.word.office;

import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.StreamOpenOfficeDocumentConverter;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.ConnectException;

/**
 * @author: Gibbons
 * @create: 2018/06/08 19:01
 * @description:
 **/
public class OfficeToPDF {
    private static OpenOfficeConnectionPool pool = new OpenOfficeConnectionPool(3);

    /**
     * @param sourceFile 源文件, 绝对路径. 可以是Office2003-2007全部格式的文档, Office2010的没测试. 包括.doc,
     *                   .docx, .xls, .xlsx, .ppt, .pptx等. 示例: F:\\office\\source.doc
     * @param destFile   目标文件. 绝对路径. 示例: F:\\pdf\\dest.pdf
     * @return 操作成功与否的提示信息. 如果返回 -1, 表示找不到源文件, 或url.properties配置错误; 如果返回 0,
     * 则表示操作成功; 返回1, 则表示转换失败
     */
    public static int office2PDF(String sourceFile, String destFile) throws FileNotFoundException {
        OpenOfficeConnection connection = pool.getConnection();
        try {
            File inputFile = new File(sourceFile);
            if (!inputFile.exists()) {
                return -1;// 找不到源文件, 则返回-1
            }

            // 如果目标路径不存在, 则新建该路径
            File outputFile = new File(destFile);
            if (!outputFile.getParentFile().exists()) {
                outputFile.getParentFile().mkdirs();
            }

            connection.connect();

            DocumentConverter converter = new StreamOpenOfficeDocumentConverter(connection);
            converter.convert(inputFile, outputFile);

            connection.disconnect();
            return 0;
        } catch (ConnectException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (connection != null) {
                    connection.disconnect();
                    connection = null;
                }
            } catch (Exception e) {
            }
        }
        return 1;
    }

    public static void main(String[] args) {
        try {
            System.out.println(office2PDF("C:\\Users\\Gibbons\\Desktop\\合同模板1.doc", "C:\\Users\\Gibbons\\Desktop\\合同模板.pdf"));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }
}
