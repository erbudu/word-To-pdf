package main;


import main.util.poi.MSWordTool;
import main.util.poi.PropertiesUtil;
import org.apache.commons.io.FileUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.*;

import static main.util.poi.FileUtil.delAllFile;
import static main.util.poi.FileZip.fileToZip;
import static main.util.poi.FileZip.unZipFiles;
import static main.util.poi.FileZip.getFileType;
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
    private static String propertiesPath = jarWholePath + "template" + File.separator + "report.properties";
    private static final String PUSH_UNFINISHED = PropertiesUtil.getValue(propertiesPath,"push_unfinished");
    private static final String PUSH_FINISHED = PropertiesUtil.getValue(propertiesPath,"push_finished");
    private static final String REPORT_UNFINISHED = PropertiesUtil.getValue(propertiesPath,"report_unfinished");
    private static final String ZIP = "zip";

    public static void main(String[] args) throws Exception {
        String sourceDirector = jarWholePath + PropertiesUtil.getValue(propertiesPath, "sourceDirector");
        System.out.println("源文件夹路径"+sourceDirector);
        while (true) {
            //连接ftp下载源文件到本地
            //downloadFtpFile(PUSH_UNFINISHED,sourceDirector);

            long old = System.currentTimeMillis();
            projectBegin(sourceDirector);
            long now = System.currentTimeMillis();
            System.out.println("共耗时：" + ((now - old) / 1000.0) + "秒");

            Thread.sleep(2000);
        }

    }


    /** 报告生成开始,解压文件
     * @param sourceDirector 源文件夹
     * @throws Exception
     */
    private static void projectBegin(String sourceDirector) throws Exception {
        //获取本地文件夹下所有文件，解压zip文件
        File file = new File(sourceDirector);
        if(file.exists()) {
            File[] fs = file.listFiles();
            if((fs != null ? fs.length : 0) != 0) {
                for (File f : fs) {
                    if (f.isFile()) {
                        //获取文件类型
                        String type = getFileType(f.getName());

                        if (ZIP.equals(type)) {
                            //获取解压后文件夹路径
                            String target = unZipFiles(f,sourceDirector);
                            //获取解压后的json文件
                            JSONObject jsonObject = null;
                            String path = target+"/detectInfo.json";
                            try {
                                String input = FileUtils.readFileToString(new File(path), "UTF-8");
                                jsonObject = JSONObject.fromObject(input);

                            } catch (Exception e) {
                                e.printStackTrace();
                            }

                            String appName = "";
                            if(jsonObject != null) {
                                //获取文件名
                                appName = jsonObject.getString("appName");

                                //生成主报告
                                getAppTestReport(jsonObject,sourceDirector+File.separator+appName+"（报告）"+File.separator);
                                //生成附件
                                getImageReport(jsonObject, target,sourceDirector+File.separator+appName+"（报告）"+File.separator);
                            } else {
                                System.out.println("json文件为空！");
                            }

                            //压缩word文件，并删除原word文件夹
                            String finalZip = sourceDirector + File.separator + appName + "（报告）";
                            String pdfName = f.getName().substring(f.getName().lastIndexOf('\\')+1, f.getName().lastIndexOf('.'));
                            fileToZip(finalZip, sourceDirector, pdfName + "（报告）");

                            File deleteFinalZip = new File(sourceDirector+File.separator+appName+"（报告）");
                            delAllFile(deleteFinalZip);

                            //删除解压后的文件夹
                            File deleteTarget = new File(target.replace("/", File.separator));
                            delAllFile(deleteTarget);

                            //上传word.zip并删除
                            File zipFile = new File(sourceDirector + File.separator + pdfName + "（报告）" + ".zip");
                            InputStream input = new FileInputStream(zipFile);
                            uploadFile(REPORT_UNFINISHED,pdfName + "（报告）" + ".zip",input);
                            input.close();
                            zipFile.delete();

                            //上传app.zip文件并删除
                            File appFile = new File(sourceDirector + File.separator + pdfName + ".zip");
                            InputStream appInput = new FileInputStream(appFile);
                            uploadFile(PUSH_FINISHED,pdfName + ".zip",appInput);
                            appInput.close();
                            appFile.delete();

                            //删除服务器上app.zip源文件
                            Map<String, Object> para = new HashMap<>(16);
                            para.put("path", PUSH_UNFINISHED);
                            para.put("name",pdfName+".zip");
                            deleteFile(para);


                            //源压缩文件移动至新文件夹(暂不需要)
                            //File director = new File(finaDirector);
                            //Path source = f.toPath();
                            //Path end = director.toPath();
                            //Files.move(source, end.resolve(source.getFileName()), REPLACE_EXISTING);
                        }
                    }
                }
            }

        }

    }

    /**生成主报告
     * @param outPath 生成报告路径
     * @throws Exception error
     */
    private static void getAppTestReport(JSONObject jsonObject,String outPath) throws Exception {

        //模板文件路径
        String reportPath = jarWholePath + PropertiesUtil.getValue(propertiesPath, "reportTemplate");
        String identifyPath = jarWholePath + PropertiesUtil.getValue(propertiesPath, "identifyTemplate");

        String lwfSignPath = jarWholePath + PropertiesUtil.getValue(propertiesPath, "lwf");
        String zjSignPath = jarWholePath + PropertiesUtil.getValue(propertiesPath, "zj");
        String sealSignPath = jarWholePath + PropertiesUtil.getValue(propertiesPath, "company");
        String identifySignPath = jarWholePath + PropertiesUtil.getValue(propertiesPath, "identify");

        //报告保存名称
        String reportName = "证据固定报告.docx";
        String identifyName = "司法鉴定报告.docx";
        String pdfName = "证据固定报告.pdf";

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
        changer.setTemplateReturnDoc("1".equals(type) ? reportPath : identifyPath);

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

        //签名插入
        Map<String,String> map4 = new HashMap<>(16);
        Map<String,String> map5 = new HashMap<>(16);
        map4.put("filepath", lwfSignPath);
        map4.put("type", "small");
        map5.put("林伟烽", "");
        changer.replaceBookMarkPhoto(map5,map4);

        map4.put("filepath", zjSignPath);
        map4.put("type", "small");
        map5.remove("林伟烽");
        map5.put("张剑", "");
        changer.replaceBookMarkPhoto(map5,map4);

        //根据不同类型选择不同公司印章
        map4.put("filepath", "1".equals(type) ? sealSignPath : identifySignPath);
        map4.put("type", "middle");
        map5.remove("张剑");
        map5.put("公章", "");
        changer.replaceBookMarkPhoto(map5,map4);

        //到服务器存档
        reportName = "1".equals(type) ? reportName : identifyName;
        changer.saveAs(outPath+reportName,outPath);
        System.out.println(appName+"文件生成成功！");

        //转换成PDF文件,删除原word文件
        //DocToPdf.doc2pdf(outPath+reportName,outPath+appName+pdfName);

    }


    /**
     * @param jsonObject json文件
     * @param imagePath 图片路径
     * @param outPath 输出路径
     * @throws Exception error
     */
    private static void getImageReport(JSONObject jsonObject,String imagePath, String outPath) throws Exception {
        //模板路径
        String reportPath = jarWholePath + PropertiesUtil.getValue(propertiesPath, "imageTemplate");

        //报告保存名称
        String reportName = "附件一 APP运行内容固定过程.docx";
        String pdfName = "附件一 APP运行内容固定过程.pdf";
        MSWordTool changer = new MSWordTool();
        XWPFDocument doc = changer.setTemplateReturnDoc(reportPath);

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
                    aHeaderPara.createRun().setText("附件一");
                } else if ((aHeaderPara.getAlignment().toString()).equals("LEFT")) {
                    aHeaderPara.createRun().setText("附件一");
                    aHeaderPara.createRun().addTab();
                    aHeaderPara.createRun().addTab();
                    aHeaderPara.createRun().setText(docNO);
                } else {
                    aHeaderPara.createRun().setText("附件一");
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
        System.out.println(appName+"文件生成成功！");

        //转换成PDF文件
        //DocToPdf.doc2pdf(outPath+reportName,outPath+appName+pdfName);

    }


}
