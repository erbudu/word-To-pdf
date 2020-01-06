package main;

import main.util.poi.MSWordTool;
import main.util.poi.PropertiesUtil;
import org.apache.commons.io.FileUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.log4j.*;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.text.SimpleDateFormat;
import java.util.*;

import static java.nio.file.StandardCopyOption.REPLACE_EXISTING;
import static main.util.poi.FileUtil.delAllFile;
import static main.util.poi.FileUtil.delFile;
import static main.util.poi.FileZip.*;
import static main.util.poi.FtpUtil.deleteFile;
import static main.util.poi.FtpUtil.downloadFtpFile;
import static main.util.poi.FtpUtil.uploadFile;

/**
 * app检测司法鉴定意见书
 *
 * @author Masai
 * @date 2019-03-07
 */
public class CreateAppReport {
    /**
     * 获取jar包绝对路径
     */
    private static String jarWholePath = CreateAppReport.class.getProtectionDomain().getCodeSource().getLocation()
            .getFile().substring(1).replace("Report.jar","");
    private static String propertiesPath = jarWholePath + File.separator + "template/report.properties";
    private static final String PUSH_UNFINISHED = PropertiesUtil.getValue(propertiesPath,"push_unfinished");
    private static final String PUSH_FINISHED = PropertiesUtil.getValue(propertiesPath,"push_finished");
    private static final String REPORT_UNFINISHED = PropertiesUtil.getValue(propertiesPath,"report_unfinished");
    private static final String ZIP = "zip";

    public static Log log = LogFactory.getLog(CreateAppReport.class);

    static {
        PropertyConfigurator.configure(jarWholePath + File.separator + "template" + File.separator
                + "log4j.properties");
    }

    public static void main(String[] args) throws Exception {
        String sourceDirector = jarWholePath + PropertiesUtil.getValue(propertiesPath, "sourceDirector");
        projectBegin(sourceDirector);
        while (true) {
            //连接ftp下载源文件到本地1
            downloadFtpFile(PUSH_UNFINISHED,sourceDirector);

            projectBegin(sourceDirector);

            Thread.sleep(3000);
        }
    }


    /** 报告生成开始,解压文件
     * @param sourceDirector 源文件夹
     * @throws Exception
     */
    private static void projectBegin(String sourceDirector) {
        //获取本地文件夹下所有文件，解压zip文件
        File file = new File(sourceDirector);
        if(file.exists()) {
            File[] fs = file.listFiles();
            //判断文件存在，且文件大于0KB
            if((fs != null ? fs.length : 0) != 0) {
                for (File f : fs) {
                    if (f.isFile() && f.length() != 0) {
                        //获取文件类型
                        String type = getFileType(f.getName());
                        log.error(f.getName());
                        if (ZIP.equals(type)) {
                            //获取解压后文件夹路径
                            String target = unZipFiles(f,sourceDirector);
                            //获取解压后的json文件
                            JSONObject jsonObject ;
                            String path = target + "/detectInfo.json";
                            try {
                                String input = FileUtils.readFileToString(new File(path), "UTF-8");
                                jsonObject = JSONObject.fromObject(input);

                            } catch (Exception e) {
                                log.error("解压包内找不到对应的json文件！");
                                File deleteTarget = new File(target.replace("/", File.separator));
                                delAllFile(deleteTarget);
                                continue;
                            }

                            String appName = "";
                            String outPath = "" ;
                            if(jsonObject != null) {
                                //获取文件名
                                appName = jsonObject.getString("appName");
                                outPath = sourceDirector + File.separator + appName + "（报告）" ;
                                //生成主报告
                                getAppTestReport(jsonObject, target, outPath + File.separator + "ViolationReport" + File.separator);
                                //生成附件
//                                getImageReport(jsonObject, target,sourceDirector+File.separator+appName+"（报告）"+File.separator);
                            }

                            //将源文件中的复制到生成word报告文件夹中，需要一一对应，文件为空的，生成空文件夹
                            String appIcon = outPath + File.separator + "AppIcon";
                            String contentViolation = outPath + File.separator + "ContentViolation";
                            String descImage = outPath + File.separator + "DescImage";
                            String packApk = outPath + File.separator + "PackApk";
                            String detectText = outPath + File.separator + "数据传输格式.txt";

                            copyFolder(target + "/icon", appIcon);
                            copyFolder(target + "/detectImg", contentViolation);
                            copyFolder(target + "/DescImage", descImage);
                            copyFolder(target + "/apk", packApk);
                            copyFile(target + "/detectInfo1.txt", detectText);


                            // 压缩word报告文件夹，并删除原word报告文件夹
                            try {
                                zipCompress(outPath, outPath+".zip");
                            } catch (Exception e) {
                                e.printStackTrace();
                            }
                            File deleteFinalZip = new File(outPath);
                            delAllFile(deleteFinalZip);

                            //删除解压后的文件夹
                            File deleteTarget = new File(target.replace("/", File.separator));
                            delAllFile(deleteTarget);

                            handleReport(sourceDirector, outPath, appName, f);

                        }
                    }
                }
            }
        }
    }


    /** 处理生成的报告文件
     * @param sourceDirector 源文件夹
     * @param outPath 目标文件夹
     * @param appName app名字
     * @param f 文件
     */
    private static void handleReport(String sourceDirector, String outPath, String appName, File f) {
        //上传word.zip并删除源文件
        try {
            File zipFile = new File(outPath + ".zip");
            InputStream input = new FileInputStream(zipFile);
            uploadFile(REPORT_UNFINISHED,appName + "（报告）.zip",input);
            input.close();
            zipFile.delete();
        } catch (IOException e) {
            e.printStackTrace();
            log.error("找不到报告zip文件！");
        }


        //上传app.zip文件
        try{
            File appFile = new File(sourceDirector + File.separator + appName + ".zip");
            InputStream appInput = new FileInputStream(appFile);
            uploadFile(PUSH_FINISHED,appName + ".zip",appInput);
            appInput.close();
//            appFile.delete();
        } catch (IOException e) {
            e.printStackTrace();
            log.error("找不到源zip文件");
        }

        //删除服务器上app.zip源文件
        Map<String, Object> para = new HashMap<>(16);
        para.put("path", PUSH_UNFINISHED);
        para.put("name",appName+".zip");
        deleteFile(para);


        //源压缩文件移动至新文件夹(本地存档)
        try {
            String savedDirector = jarWholePath + PropertiesUtil.getValue(propertiesPath, "savedDirector");
            File director = new File(savedDirector);
            Path source = f.toPath();
            Path end = director.toPath();
            Files.move(source, end.resolve(source.getFileName()), REPLACE_EXISTING);
            log.info(appName+"报告压缩包已生成");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    /**生成主报告
     * @param jsonObject json数据
     * @param imagePath 图片路径
     * @param outPath 输出路径
     */
    private static void getAppTestReport(JSONObject jsonObject, String imagePath, String outPath) {

        //模板文件路径
        String reportPath = jarWholePath + PropertiesUtil.getValue(propertiesPath, "reportTemplate");
        String identifyPath = jarWholePath + PropertiesUtil.getValue(propertiesPath, "identifyTemplate");

        String sealSignPath = jarWholePath + PropertiesUtil.getValue(propertiesPath, "company");
        String identifySignPath = jarWholePath + PropertiesUtil.getValue(propertiesPath, "identify");

        //报告保存名称
        String reportName = "证据固定报告.docx";
        String identifyName = "司法鉴定意见书.docx";
        String reportPdfName = "证据固定报告.pdf";
        String identifyPdfName = "司法鉴定意见书.pdf";

        //获取当前时间
        SimpleDateFormat date = new SimpleDateFormat("yyyy年MM月dd日");
        Date now = new Date();
        String createTime = date.format(now);
        //获取参数
        String acceptData = jsonObject.getString("acceptData");
        String appIP = jsonObject.getString("appIP");
        String appName = jsonObject.getString("appName");
        String basicCase = jsonObject.getString("basicCase");
        String client = jsonObject.getString("client");
        String docNO = jsonObject.getString("docNO");
        String endTime = jsonObject.getString("endTime");
        String fileMD5 = jsonObject.getString("fileMD5");
        String fileName = jsonObject.getString("fileName");
        String matter = jsonObject.getString("matter");
        String fileSize = jsonObject.getString("fileSize");
        String sourceURL = jsonObject.getString("sourceURL");
        String startTime = jsonObject.getString("startTime");
        String type = jsonObject.get("type")+"";

        MSWordTool changer = new MSWordTool();
        try {
            changer.setTemplateReturnDoc("1".equals(type) ? reportPath : identifyPath);
        } catch (Exception e) {
            e.printStackTrace();
            log.info("模板报告获取失败，请检查模板文件路径");
        }

        Map<String,String> map = new HashMap<>(16);
        map.put("备案编号", docNO);
        changer.replaceBookMarkText(map,false,false,10,"黑体");

        Map<String,String> map2 = new HashMap<>(16);
        map2.put("委托人", client);
        map2.put("固证事项", matter);
        map2.put("受理日期", acceptData);
        map2.put("基本案情", basicCase);
        map2.put("鉴定时间", startTime+"至"+endTime);
        map2.put("鉴定过程时间", startTime+"至"+endTime);
        map2.put("报告日期", createTime);
        map2.put("APP程序名2", appName);
        changer.replaceBookMarkText(map2,false,false,14,"仿宋_GB2312");

        List<Map<String,String>> summaryMapList = new ArrayList<>();
        Map<String,String> map3 = new HashMap<>(16);
        map3.put("来源网址", sourceURL);
        map3.put("应用名称", appName);
        map3.put("安装包文件名", fileName);
        map3.put("大小", fileSize);
        map3.put("MD5", fileMD5);
        summaryMapList.add(map3);
        changer.fillTableAtBookMark("资料摘要", summaryMapList);

        map3.remove("来源网址", sourceURL);
        changer.fillTableAtBookMark("分析说明", summaryMapList);

        Map<String,String> map6 = new HashMap<>(16);
        map6.put("APP程序名", appName);
        map6.put("IP地址", appIP);
        changer.replaceBookMarkText(map6,false,false,12,"仿宋_GB2312");

        //根据不同类型选择不同公司印章
        Map<String,String> map4 = new HashMap<>(16);
        Map<String,String> map5 = new HashMap<>(16);
        map4.put("filepath", "1".equals(type) ? sealSignPath : identifySignPath);
        map4.put("type", "small");
        map5.put("公章", "");
        changer.replaceBookMarkPhoto(map5,map4);

        //获取图片json数据
        JSONArray imgArray = jsonObject.getJSONArray("imgSet");

        //附录一软件名
        Map<String,String> map7 = new HashMap<>(16);
        map7.put("软件名", "“"+appName+"”");
        changer.replaceBookMarkText(map7,false,false,10,"仿宋_GB2312");

        //插入图片一
        Map<String,String> map8 = new HashMap<>(16);
        Map<String,String> map9 = new HashMap<>(16);
        map8.put("filepath", imagePath+imgArray.getJSONObject(0).getString("img"));
        map8.put("type", "middle");
        map9.put("图片一", "");
        changer.replaceBookMarkPhoto(map9,map8);

        // 批量插入图片
        changer.fillPictureTableAtBookMark("图片表格",imagePath,imgArray);

        //到服务器存档
        reportName = "1".equals(type) ? reportName : identifyName;
        changer.saveAs(outPath + docNO + appName + reportName, outPath);

        //转换成PDF文件,删除原word文件
        String pdfName = "1".equals(type) ? reportPdfName : identifyPdfName;
        DocToPdf.doc2pdf(outPath + docNO + appName + reportName, outPath + docNO + appName + pdfName);

    }


    /** 输出图片文档
     * @param jsonObject json文件
     * @param imagePath 图片路径
     * @param outPath 输出路径
     */
    private static void getImageReport(JSONObject jsonObject,String imagePath, String outPath) {
        //模板路径
        String reportPath = jarWholePath + PropertiesUtil.getValue(propertiesPath, "imageTemplate");

        //报告保存名称
        String reportName = "附件 APP运行内容固定过程.docx";
        String pdfName = "附件 APP运行内容固定过程.pdf";
        MSWordTool changer = new MSWordTool();
        XWPFDocument doc = null;
        try {
            doc = changer.setTemplateReturnDoc(reportPath);
        } catch (Exception e) {
            e.printStackTrace();
            log.info("模板附件获取失败，请检查模板文件路径");
        }

        //获取json数据
        JSONArray imgArray = jsonObject.getJSONArray("imgSet");
        String docNO = jsonObject.get("docNO").toString();
        String appName = jsonObject.get("appName").toString();

        Map<String,String> map = new HashMap<>(16);
        Map<String,String> map1 = new HashMap<>(16);
        Map<String,String> map2 = new HashMap<>(16);

        //标题软件名
        map1.put("软件名", "“"+appName+"”");
        changer.replaceBookMarkText(map1,false,false,10,"仿宋_GB2312");

        //页眉备案编号
        List<XWPFHeader> pageHeaders = doc.getHeaderList();
        for (XWPFHeader pageHeader : pageHeaders) {
            List<XWPFParagraph> headerPara = pageHeader.getParagraphs();
            for (XWPFParagraph aHeaderPara : headerPara) {
                if ((aHeaderPara.getAlignment().toString()).equals("RIGHT")) {
                    aHeaderPara.createRun().setText(docNO);
                    aHeaderPara.createRun().addTab();
                    aHeaderPara.createRun().addTab();
                    aHeaderPara.createRun().setText("附件");
                } else if ((aHeaderPara.getAlignment().toString()).equals("LEFT")) {
                    aHeaderPara.createRun().setText("附件");
                    aHeaderPara.createRun().addTab();
                    aHeaderPara.createRun().addTab();
                    aHeaderPara.createRun().setText(docNO);
                } else {
                    aHeaderPara.createRun().setText("附件");
                    aHeaderPara.createRun().addTab();
                    aHeaderPara.createRun().addTab();
                    aHeaderPara.createRun().setText(docNO);
                }
                aHeaderPara.removeRun(0);
            }
        }

        //插入图片一
        map.put("filepath", imagePath+imgArray.getJSONObject(0).getString("img"));
        map.put("type", "big");
        map2.put("图片一", "");
        changer.replaceBookMarkPhoto(map2,map);

        //在表格中插入剩下的图片
        changer.fillPictureTableAtBookMark("图片表格",imagePath,imgArray);

        //到服务器存档
        changer.saveAs(outPath+reportName,outPath);

        //转换成PDF文件
        //DocToPdf.doc2pdf(outPath+reportName,outPath+appName+pdfName);

    }

    /**复制文件夹所有文件到新文件夹
     * @param strOldPath 源文件夹
     * @param strNewPath 新文件夹
     */
    private static void copyFolder(String strOldPath, String strNewPath) {
        File fOldFolder = new File(strOldPath);
        try {
            File fNewFolder = new File(strNewPath);
            if (!fNewFolder.exists()) {
                //不存在就创建一个文件夹
                fNewFolder.mkdirs();
            }

            //获取旧文件夹里面所有的文件
            if (fOldFolder.exists()) {
                File[] arrFiles = fOldFolder.listFiles();
                for (int i = 0; i < arrFiles.length; i++) {
                    if (arrFiles[i].isDirectory()) {
                        copyFolder(strOldPath + File.separator + arrFiles[i].getName(), strNewPath + File.separator + arrFiles[i].getName());
                    } else {
                        copyFile(strOldPath + File.separator + arrFiles[i].getName(), strNewPath + File.separator + arrFiles[i].getName());
                    }
                }
            }
        }
        catch (Exception e) {
            log.error("复制文件失败：" + strOldPath);
        }
    }

    private static void copyFile(String strOldPath, String strNewPath) {
        try {
            File fOldFile = new File(strOldPath);
            if (fOldFile.exists()) {
                int byteread = 0;
                InputStream inputStream = new FileInputStream(fOldFile);
                FileOutputStream fileOutputStream = new FileOutputStream(strNewPath);
                byte[] buffer = new byte[1444];
                while ( (byteread = inputStream.read(buffer)) != -1) {
                    //三个参数，第一个参数是写的内容，第二个参数是从什么地方开始写，第三个参数是需要写的大小
                    fileOutputStream.write(buffer, 0, byteread);
                }
                inputStream.close();
                fileOutputStream.close();
            }
        }
        catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            System.out.println("复制单个文件出错");
            e.printStackTrace();
        }
    }
}
