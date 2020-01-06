package main;

import com.aspose.words.Document;
import com.aspose.words.License;
import com.aspose.words.SaveFormat;
import main.util.poi.PropertiesUtil;

import java.io.*;

import static main.CreateAppReport.log;
import static main.util.poi.FileUtil.delFile;

/**
 * doc转PDF
 * @author Masai
 * @date 2019-03-11
 */

class DocToPdf {
    private static String jarWholePath = CreateAppReport.class.getProtectionDomain().getCodeSource().getLocation()
            .getFile().substring(1).replace("Report.jar","");
    private static String propertiesPath = jarWholePath + File.separator + "template/report.properties";

    private static boolean getLicense() {
        boolean result = false;
        try {
            //获取license文件
            String licensePath = jarWholePath + PropertiesUtil.getValue(propertiesPath, "licensePath");
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

    static void doc2pdf(String inPath,String outPath) {
        try {
            if (!getLicense()) {
                // 验证License 若不验证则转化出的pdf文档会有水印产生
                throw new Exception("com.aspose.words license ERROR!");
            }
        } catch (Exception e) {
            log.error("com.aspose.words license ERROR!");
            e.printStackTrace();
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
            //delFile(inPath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
