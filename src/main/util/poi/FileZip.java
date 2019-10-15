package main.util.poi;

import java.io.*;
import java.nio.charset.Charset;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

import static main.CreateAppReport.log;

/**java压缩成zip
 * @author Se7en
 */
public class FileZip {
	/**
	 * @param inputFileName 你要压缩的文件夹(整个完整路径)
	 * @param zipFileName   压缩后的文件(整个完整路径)
	 * @throws Exception
	 */
	public static Boolean zip(String inputFileName, String zipFileName) throws Exception {
		zip(zipFileName, new File(inputFileName));
		return true;
	}

	public static void zip(String zipFileName, File inputFile) throws Exception {
		ZipOutputStream out = new ZipOutputStream(new FileOutputStream(zipFileName));
		zip(out, inputFile, "");
		out.flush();
		out.close();
	}

	private static void zip(ZipOutputStream out, File f, String base) throws Exception {
		if (f.isDirectory()) {
			File[] fl = f.listFiles();
			out.putNextEntry(new ZipEntry(base + "/"));
			base = base.length() == 0 ? "" : base + "/";
			for (int i = 0; i < fl.length; i++) {
				zip(out, fl[i], base + fl[i].getName());
			}

		} else {
			out.putNextEntry(new ZipEntry(base));
			FileInputStream in = new FileInputStream(f);
			int b;
			while ((b = in.read()) != -1) {
				out.write(b);
			}
			in.close();
		}

	}

	/**解压
	 * @param zipFile
	 * @param descDir
	 */
	public static String unZipFiles(File zipFile, String descDir) {
		//解决中文文件夹乱码
		ZipFile zip = null;
		String name = "";
		try {
			zip = new ZipFile(zipFile, Charset.forName("GBK"));
			//获取压缩文件文件名
			name = zip.getName().substring(zip.getName().lastIndexOf('\\')+1, zip.getName().lastIndexOf('.'));
		} catch (IOException e) {
			e.printStackTrace();
			log.error("zip压缩包文件异常，无法解压！");
		}

		//解压后的文件路径
		String finaPath = descDir + File.separator + name;
		File pathFile = new File(finaPath);
		if (!pathFile.exists()) {
			pathFile.mkdirs();
		}

		for (Enumeration<? extends ZipEntry> entries = zip.entries(); entries.hasMoreElements(); ) {
			ZipEntry entry = entries.nextElement();
			String zipEntryName = entry.getName();
			InputStream in = null;
			try {
				in = zip.getInputStream(entry);
			} catch (IOException e) {
				e.printStackTrace();
			}
			String outPath = (descDir + File.separator + name + "/" + zipEntryName).replaceAll("\\*", "/");

			// 判断路径是否存在,不存在则创建文件路径
			File file = new File(outPath.substring(0, outPath.lastIndexOf('/')));
			if (!file.exists()) {
				file.mkdirs();
			}
			// 判断文件全路径是否为文件夹,如果是上面已经上传,不需要解压
			if (new File(outPath).isDirectory()) {
				continue;
			}

			FileOutputStream out = null;
			try {
				out = new FileOutputStream(outPath);
			} catch (FileNotFoundException e) {
				e.printStackTrace();
				log.error("找不到解压文件输出路径");
				continue;
			}
			byte[] buf1 = new byte[1024];
			int len;
			try {
				while ((len = in.read(buf1)) > 0) {
					out.write(buf1, 0, len);
				}
				in.close();
				out.close();
			} catch (IOException e) {
				e.printStackTrace();
			}

		}
		//不加入close，原压缩文件删不掉
		try {
			zip.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		log.info(name+"已解压");
		return finaPath;
	}


	/**
	 * 将存放在sourceFilePath目录下的源文件,打包成fileName名称的ZIP文件,并存放到zipFilePath。
	 * @param sourceFilePath 待压缩的文件路径
	 * @param zipFilePath 压缩后存放路径
	 * @param fileName 压缩后文件的名称
	 * @return flag
	 */
	public static boolean fileToZip(String sourceFilePath,String zipFilePath,String fileName) {
		boolean flag = false;
		File sourceFile = new File(sourceFilePath);
		if(!sourceFile.exists()) {
			log.error("待压缩的文件目录：" + sourceFilePath + " 不存在");
			flag = false;
			return flag;
		} else {
			try {
				File zipFile = new File(zipFilePath + "/" + fileName + ".zip");
				if(zipFile.exists()) {
					log.error(zipFilePath + " 目录下存在名字为：" + fileName + ".zip" + " 打包文件");
				} else {
					File[] sourceFiles = sourceFile.listFiles();
					if(null == sourceFiles || sourceFiles.length < 1) {
						log.info("待压缩的文件目录：" + sourceFilePath + " 里面不存在文件,无需压缩");
						flag = false;
						return flag;
					} else {
						ZipOutputStream zos = new ZipOutputStream(new BufferedOutputStream(new FileOutputStream(zipFile)));
						//缓冲块
						byte[] bufs = new byte[1024*10];
						for(int i=0;i<sourceFiles.length;i++) {
							ZipEntry zipEntry = new ZipEntry(sourceFiles[i].getName());
							zos.putNextEntry(zipEntry);
							BufferedInputStream bis = new BufferedInputStream(new FileInputStream(sourceFiles[i]),1024*10);
							int read = 0;
							while((read=(bis.read(bufs, 0, 1024*10))) != -1) {
								zos.write(bufs, 0, read);
							}
							//关闭
							if(null != bis) {
								bis.close();
							}
						}
						flag = true;
						//关闭
						if(null != zos) {
							zos.close();
						}
					}
				}

			} catch (IOException e) {
				e.printStackTrace();
				throw new RuntimeException(e);
			}
		}
		return flag;
	}

	/**获取文件后缀名
	 * @param fileName
	 * @return
	 */
	public static String getFileType(String fileName) {
		String[] strArray = fileName.split("\\.");
		int suffixIndex = strArray.length -1;
		return strArray[suffixIndex];
	}

}