package main.util.poi;

import main.CreateAppReport;
import org.apache.commons.net.ftp.FTP;
import org.apache.commons.net.ftp.FTPClient;
import org.apache.commons.net.ftp.FTPFile;
import org.apache.commons.net.ftp.FTPReply;

import java.io.*;
import java.net.SocketException;
import java.util.Map;

/**
 * ftp上传、下载、删除文件
 *
 * @author Masai
 * @date 2019-03-15
 */

public class FtpUtil {

    private static FTPClient CLIENT = null;
    private static String LOCAL_CHARSET = "GBK";
    private static String SERVER_CHARSET = "iso-8859-1";
    private static String jarWholePath = CreateAppReport.class.getProtectionDomain().getCodeSource().getLocation()
            .getFile().substring(1).replace("Report.jar","");
    private static String propertiesPath = jarWholePath + "template" + File.separator + "report.properties";
    private static final String FTP_HOST = PropertiesUtil.getValue(propertiesPath,"ftpHost");
    private static final String FTP_USERNAME = PropertiesUtil.getValue(propertiesPath,"ftpUserName");
    private static final String FTP_PASSWORD = PropertiesUtil.getValue(propertiesPath,"ftpPassWord");

    /**
     * 获取FTPClient对象
     *
     *ftpHost  FTP主机服务器
     *ftpPort  FTP端口 默认为21
     * @return
     */
    public static void getFTPClient() {
        try {
            CLIENT = new FTPClient();
            // 连接FTP服务器
            CLIENT.connect(FTP_HOST, 21);
            // 登陆FTP服务器
            CLIENT.login(FTP_USERNAME, FTP_PASSWORD);
            if (!FTPReply.isPositiveCompletion(CLIENT.getReplyCode())) {
                System.out.println("未连接到FTP，用户名或密码错误。");
                CLIENT.disconnect();
            } else {
                int reply;
                reply = CLIENT.getReplyCode();
                if (!FTPReply.isPositiveCompletion(reply)) {
                    CLIENT.disconnect();
                }

                CLIENT.setControlEncoding("GBK");
                CLIENT.setFileType(FTPClient.BINARY_FILE_TYPE);
                CLIENT.enterLocalPassiveMode();
            }
        } catch (SocketException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**从FTP服务器下载文件
     * @param ftpPath FTP服务器中文件所在路径 格式： ftptest/aa
     * @param localPath 下载到本地的位置 格式：H:/download
     * @return true
     */
    public static boolean downloadFtpFile(String ftpPath, String localPath) {
        boolean isAppend = false;
        try {
            getFTPClient();
            CLIENT.changeWorkingDirectory(ftpPath);
            //获取ftpPath文件夹下所有文件
            FTPFile[] ftpFiles = CLIENT.listFiles(ftpPath);
            System.out.println("文件个数:"+ftpFiles.length);

            for (FTPFile file : ftpFiles) {
                String name = file.getName();
                File localFile = new File(localPath + File.separatorChar + name);

                if(!new File(localPath).exists()){
                    new File(localPath).mkdirs();
                }

                OutputStream os = new FileOutputStream(localFile);
                isAppend =  CLIENT.retrieveFile(new String(name.getBytes(LOCAL_CHARSET),SERVER_CHARSET), os);
                os.close();

            }
            CLIENT.logout();
            return isAppend;

        } catch (FileNotFoundException | SocketException e) {
            System.out.println("没有找到文件");
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
            e.printStackTrace();
        }
        return isAppend;
    }

    /**
     * Description: 向FTP服务器上传文件
     * @param ftpPath  FTP服务器中文件所在路径 格式： ftptest/aa
     * @param fileName ftp文件名称
     * @param input 文件流
     * @return 成功返回true，否则返回false
     */
    public static boolean uploadFile(String ftpPath,String fileName,InputStream input) {
        boolean success = false;
        try {
            getFTPClient();
            CLIENT.changeWorkingDirectory(ftpPath);
            fileName = new String(fileName.getBytes(LOCAL_CHARSET),SERVER_CHARSET);
            success = CLIENT.storeFile(fileName, input);
            input.close();
            CLIENT.logout();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (CLIENT.isConnected()) {
                try {
                    CLIENT.disconnect();
                } catch (IOException ioe) {
                }
            }
        }
        return success;
    }

    /**
     * Description: 向FTP服务器上传图片
     * @param ftpPath  FTP服务器中文件所在路径,把路径按照级别放到数组中
     * @param fileName ftp文件名称
     * @param input 文件流
     * @return 成功返回true，否则返回false
     */
    public static boolean uploadFile(String[] ftpPath,String fileName,InputStream input) {
        boolean success = false;
        try {
            getFTPClient();
            String path = "";
            for(int i=0;i<ftpPath.length;i++){
                path = new String(ftpPath[i].getBytes(LOCAL_CHARSET),SERVER_CHARSET);
                if(i == 1){
                    CLIENT.removeDirectory(path);
                }
                CLIENT.makeDirectory(path);
                CLIENT.changeWorkingDirectory(path);
            }
            FTPFile[] files = CLIENT.listFiles();
            for(FTPFile ff:files){
                CLIENT.deleteFile(ff.getName());
            }
            fileName = new String(fileName.getBytes(LOCAL_CHARSET),SERVER_CHARSET);
            success = CLIENT.storeFile(fileName, input);
            input.close();
            CLIENT.logout();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (CLIENT.isConnected()) {
                try {
                    CLIENT.disconnect();
                } catch (IOException ioe) {
                }
            }
        }
        return success;
    }

    /**
     * @MethodName  deleteFile
     * @Description 删除ftp文件夹下面的视频文件
     * @param para
     * @return
     *
     */
    public static boolean deleteFile(Map<String, Object> para){
        boolean isAppend = false;
        try {
            String path = para.get("path")+"";
            String name = para.get("name")+"";
            getFTPClient();
            CLIENT.changeWorkingDirectory(path);
            name = new String(name.getBytes(LOCAL_CHARSET),SERVER_CHARSET);
            CLIENT.dele(name);
            CLIENT.logout();
            isAppend = true;
        } catch (Exception e) {
            e.printStackTrace();
        }finally {
            if (CLIENT.isConnected()) {
                try {
                    CLIENT.disconnect();
                } catch (IOException ioe) {
                }
            }
        }
        return isAppend;
    }

    /**
     * @MethodName  deleteFolder
     * @Description 删除ftp文件夹
     * @param folder
     * @return
     *
     */
    public static boolean deleteFolder(String folder){
        boolean isAppend = false;
        try {
            getFTPClient();
            folder = new String(folder.getBytes(LOCAL_CHARSET),SERVER_CHARSET);
            isAppend = CLIENT.removeDirectory(folder);
            CLIENT.logout();
        } catch (Exception e) {
            e.printStackTrace();
        }finally {
            if (CLIENT.isConnected()) {
                try {
                    CLIENT.disconnect();
                } catch (IOException ioe) {
                }
            }
        }
        return isAppend;
    }

    /**
     * @MethodName  reName
     * @Description 修改文件名称
     * @param para
     * @return
     *
     */
    public static boolean reName(Map<String, Object>para){
        boolean isAppend = false;
        try {
            String path = para.get("path")+"";
            String name = para.get("name")+"";
            getFTPClient();
            path = new String(path.getBytes(LOCAL_CHARSET),SERVER_CHARSET);
            name = new String(name.getBytes(LOCAL_CHARSET),SERVER_CHARSET);
            isAppend = CLIENT.rename(path, name);
            CLIENT.logout();
        } catch (Exception e) {
            e.printStackTrace();
        }finally {
            if (CLIENT.isConnected()) {
                try {
                    CLIENT.disconnect();
                } catch (IOException ioe) {
                }
            }
        }
        return isAppend;
    }


}
