package main;

import com.aspose.words.Document;
import com.aspose.words.License;
import com.aspose.words.SaveFormat;
import main.util.poi.PropertiesUtil;

import java.io.*;

import static main.util.poi.FileUtil.delFile;

/**
 * doc转PDF
 * @author Masai
 * @date 2019-03-11
 */

class DocToPdf {

    private static boolean getLicense() {
        boolean result = false;
        try {
            //获取license文件
            String licensePath = PropertiesUtil.getValue("report.properties", "licensePath");
            InputStream inputStream = new FileInputStream(licensePath);

            License aposeLic = new License();
            aposeLic.setLicense(inputStream);
            result = true;
            inputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

    static void doc2pdf(String inPath,String outPath) throws Exception {
        if (!getLicense()) {
            // 验证License 若不验证则转化出的pdf文档会有水印产生
            throw new Exception("com.aspose.words license ERROR!");
        }

        try {
            //新建一个空白pdf文档
            File file = new File(outPath);

            FileOutputStream os = new FileOutputStream(file);
            // 支持RTF HTML,OpenDocument, PDF,EPUB, XPS转换
            Document doc = new Document(inPath);
            try {
                doc.save(os, SaveFormat.PDF);
                os.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
            //删除源文件
            delFile(inPath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
