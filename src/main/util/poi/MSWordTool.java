package main.util.poi;

import java.awt.image.BufferedImage;
import java.io.*;

import java.math.BigInteger;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import net.sf.json.JSONArray;
import org.apache.poi.POIXMLDocument;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.w3c.dom.Node;

import javax.imageio.ImageIO;

/**
 * 使用POI,进行Word相关的操作
 *
 * @author xuyu
 *         <p>
 *         Modification History:
 *         </p>
 *         <p>
 *         Date Author Description
 *         </p>

 */
public class MSWordTool {

	/** 内部使用的文档对象 **/
	private XWPFDocument document;

	private BookMarks bookMarks = null;
	PoiWordUtil poiWordUtil = new PoiWordUtil();
	/**
	 * 为文档设置模板
	 * 
	 * @param templatePath
	 *            模板文件名称
	 */
	public void setTemplate(String templatePath) {
		try {
			this.document = new XWPFDocument(POIXMLDocument.openPackage(templatePath));

			bookMarks = new BookMarks(document);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	/**
	 * 为文档设置模板
	 * 
	 * @param templatePath
	 *            模板文件名称
	 * @throws Exception 
	 */
	public XWPFDocument setTemplateReturnDoc(String templatePath) throws Exception {
		try {
			this.document = new XWPFDocument(
					POIXMLDocument.openPackage(templatePath));
			bookMarks = new BookMarks(document);

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return this.document;
	}

	/**
	 * 进行标签替换的例子,传入的Map中，key表示标签名称，value是替换的信息
	 *@param paragraphAlignment
	 * @param indicator
	 */
	public void replaceBookMark(Map<String, String> indicator,ParagraphAlignment paragraphAlignment) {
		// 循环进行替换
		Iterator<String> bookMarkIter = bookMarks.getNameIterator();
		while (bookMarkIter.hasNext()) {
			String bookMarkName = bookMarkIter.next();

			// 得到标签名称
			BookMark bookMark = bookMarks.getBookmark(bookMarkName);

			// 进行替换
			if (indicator.get(bookMarkName) != null) {
				bookMark.insertTextAtBookMark(indicator.get(bookMarkName),
						BookMark.INSERT_BEFORE,paragraphAlignment);
			}
		}
	}

	/**
	 * 根据书签，文本替换,传入的Map中，key表示标签名称，value是替换的信息
	 * 可以对文本格式进行设置
	 * @param indicator
	 */
	public void replaceBookMarkText(Map<String, String> indicator,boolean isBold,boolean isUnderline,int fontSize,String fontFamily) {
		// 循环进行替换
		Iterator<String> bookMarkIter = bookMarks.getNameIterator();
		while (bookMarkIter.hasNext()) {
			String bookMarkName = bookMarkIter.next();

			// 得到标签名称
			BookMark bookMark = bookMarks.getBookmark(bookMarkName);

			// 进行替换
			if (indicator.get(bookMarkName) != null) {
				bookMark.repalceTextAtBookMark(indicator.get(bookMarkName),isBold,isUnderline,fontSize,fontFamily);
			}
		}
	}
	
	/**
	 * 根据书签合并单元格
	 * 
	 * @param bookMarkName
	 */
	public void mergeBookMarkTable(String bookMarkName) {
		// 得到标签名称
		BookMark bookMark = bookMarks.getBookmark(bookMarkName);
		XWPFTable table = bookMark.getContainerTable();
		poiWordUtil.mergeCellsHorizontal(table, 1, 1, 3);
	}

	public void replaceText(Map<String, String> bookmarkMap, String bookMarkName) {

		// 首先得到标签
		BookMark bookMark = bookMarks.getBookmark(bookMarkName);
		// 获得书签标记的表格
		XWPFTable table = bookMark.getContainerTable();
		// 获得所有的表
		// Iterator<XWPFTable> it = document.getTablesIterator();

		if (table != null) {
			// 得到该表的所有行
			int rcount = table.getNumberOfRows();
			for (int i = 0; i < rcount; i++) {
				XWPFTableRow row = table.getRow(i);

				// 获到改行的所有单元格
				List<XWPFTableCell> cells = row.getTableCells();
				for (XWPFTableCell c : cells) {
					for (Entry<String, String> e : bookmarkMap.entrySet()) {
						if (c.getText().equals(e.getKey())) {

							// 删掉单元格内容
							c.removeParagraph(0);

							// 给单元格赋值
							c.setText(e.getValue());
						}
					}
				}
			}
		}
	}

	/**
	 * 循环生成表格多条信息
	 * 针对只在表格后插入新行的情况
	 * @param bookMarkName
	 * @param content
	 */
	public void fillTableAtBookMark(String bookMarkName, List<Map<String, String>> content) {

		// rowNum来比较标签在表格的哪一行
		int rowNum = 0;

		// 首先得到标签
		BookMark bookMark = bookMarks.getBookmark(bookMarkName);
		Map<String, String> columnMap = new HashMap<>(16);
		Map<String, Node> styleNode = new HashMap<>(16);

		// 标签是否处于表格内
		if (bookMark.isInTable()) {

			// 获得标签对应的Table对象和Row对象
			XWPFTable table = bookMark.getContainerTable();

			XWPFTableRow row = bookMark.getContainerTableRow();
			CTRow ctRow = row.getCtRow();
			List<XWPFTableCell> rowCell = row.getTableCells();
			for (int i = 0; i < rowCell.size(); i++) {
				columnMap.put(i + "", rowCell.get(i).getText().trim());

				// 获取该单元格段落的xml，得到根节点
				Node node1 = rowCell.get(i).getParagraphs().get(0).getCTP().getDomNode();

				// 遍历根节点的所有子节点
				for (int x = 0; x < node1.getChildNodes().getLength(); x++) {
					if (node1.getChildNodes().item(x).getNodeName()
							.equals(BookMark.RUN_NODE_NAME)) {
						Node node2 = node1.getChildNodes().item(x);

						// 遍历所有节点为"w:r"的所有自己点，找到节点名为"w:rPr"的节点
						for (int y = 0; y < node2.getChildNodes().getLength(); y++) {
							if (node2.getChildNodes().item(y).getNodeName()
									.endsWith(BookMark.STYLE_NODE_NAME)) {

								// 将节点为"w:rPr"的节点(字体格式)存到HashMap中
								styleNode.put(i + "", node2.getChildNodes().item(y));
							}
						}
					} else {
						continue;
					}
				}
			}

			// 循环对比，找到该行所处的位置，删除改行
			for (int i = 0; i < table.getNumberOfRows(); i++) {
				if (table.getRow(i).equals(row)) {
					rowNum = i;
					break;
				}
			}
			table.removeRow(rowNum);
			for (int i = 0; i < content.size(); i++) {
				// 创建新的一行,单元格数是表的第一行的单元格数,后面添加数据时，要判断单元格数是否一致
				XWPFTableRow tableRow = table.createRow();
				CTTrPr trPr = tableRow.getCtRow().addNewTrPr();
				CTHeight ht = trPr.addNewTrHeight();
				ht.setVal(BigInteger.valueOf(360));
			}

			// 得到表格行数
			int rcount = table.getNumberOfRows();
			for (int i = rowNum; i < rowNum+content.size(); i++) {
				XWPFTableRow newRow = table.getRow(i);
				// 判断newRow的单元格数是不是该书签所在行的单元格数
				if (newRow.getTableCells().size() != rowCell.size()) {

					// 计算newRow和书签所在行单元格数差的绝对值
					// 如果newRow的单元格数多于书签所在行的单元格数，不能通过此方法来处理，可以通过表格中文本的替换来完成
					// 如果newRow的单元格数少于书签所在行的单元格数，要将少的单元格补上
					int sub = Math.abs(newRow.getTableCells().size()
							- rowCell.size());
					// 将缺少的单元格补上
					for (int j = 0; j < sub; j++) {
						newRow.addNewTableCell();
					}
				}

				List<XWPFTableCell> cells = newRow.getTableCells();

				for (int j = 0; j < cells.size(); j++) {
					XWPFParagraph para = cells.get(j).getParagraphs().get(0);
					para.setAlignment(ParagraphAlignment.CENTER);
					//设置段落高度
					PoiWordUtil.setParHeight(para,"360");
					/************文字在表格内居中start**************/
					CTTc cttc = cells.get(j).getCTTc();
					CTTcPr ctPr = cttc.addNewTcPr();
					ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
					cttc.getPArray(0).addNewPPr().addNewJc().setVal(STJc.CENTER);

					if (content.get(i - rowNum).get(columnMap.get(j + "")) != null) {
						XWPFRun run = para.insertNewRun(0);
						// 改变单元格的值，标题栏不用改变单元格的值
						run.setText(content.get(i - rowNum).get(columnMap.get(j + ""))+ "");

						// 将单元格段落的字体格式设为原来单元格的字体格式
						run.getCTR().getDomNode().insertBefore(styleNode.get(j + "").cloneNode(true),
								run.getCTR().getDomNode().getFirstChild());
					}
				}
			}
			if(content.size()==0){
				table.createRow();
			}
		}
	}


	/**
	 * 循环插入图片到表格中
	 * @param bookMarkName
	 * @param imgArray
	 */
	public void fillPictureTableAtBookMark(String bookMarkName, String imagePath, JSONArray imgArray) {
		// 首先得到标签
		BookMark bookMark = bookMarks.getBookmark(bookMarkName);
		Map<String, String> columnMap = new HashMap<>(16);
		Map<String, Node> styleNode = new HashMap<>(16);

		// 标签是否处于表格内
		if (bookMark.isInTable()) {

			// 获得标签对应的Table对象和Row对象
			XWPFTable table = bookMark.getContainerTable();

			int size = imgArray.size();
			int rowNumber = size%2 == 0 ? (size/2) : (size/2)+1;

			for (int i = 0; i < rowNumber; i++) {
				// 创建新的一行,单元格数是表的第一行的单元格数
				XWPFTableRow tableRow = table.createRow();
				CTTrPr trPr = tableRow.getCtRow().addNewTrPr();
				CTHeight ht = trPr.addNewTrHeight();
				ht.setVal(BigInteger.valueOf(360));
			}

			//已新建完空行，循环插入数据
			int k = 1;
			for (int i = 0; i < rowNumber; i++) {
				XWPFTableRow newRow = table.getRow(i);
				List<XWPFTableCell> cells = newRow.getTableCells();
				for (int j = 0; j < cells.size(); j++) {
					//表格样式，上对齐
					CTTc cttc = cells.get(j).getCTTc();
					CTTcPr ctPr = cttc.addNewTcPr();
					ctPr.addNewVAlign().setVal(STVerticalJc.TOP);

					XWPFParagraph para = cells.get(j).getParagraphs().get(0);
					XWPFRun run = para.createRun();
					if(k < imgArray.size()) {
						String imgFilePath = imagePath + imgArray.getJSONObject(k).getString("img");

						try {
							FileInputStream fileInputImg = new FileInputStream(imgFilePath);
							BufferedImage bufferedImg = ImageIO.read(fileInputImg);
							int imgWidth = bufferedImg.getWidth();
							int imgHeight = bufferedImg.getHeight();
							fileInputImg.close();
							fileInputImg = new FileInputStream(imgFilePath);

							run.addPicture(fileInputImg,Document.PICTURE_TYPE_JPEG,
									bookMarkName, Units.toEMU(210),Units.toEMU(((float)210/imgWidth)*imgHeight));
							fileInputImg.close();
						} catch (Exception e) {
							e.printStackTrace();
						}
						k++;
					}

				}
			}

		}
	}


	/**
	 * 循环生成表格多条信息
	 * 正对需要在表格中间插入新行的情况
	 * @param bookMarkName 书签
	 * @param content value
	 */
	public void addTableRowAtBookMark(String bookMarkName, List<Map<String, String>> content) {

		// rowNum来比较标签在表格的哪一行
		int rowNum = 0;

		// 首先得到标签
		BookMark bookMark = bookMarks.getBookmark(bookMarkName);
		Map<String, String> columnMap = new HashMap<>(16);
		Map<String, Node> styleNode = new HashMap<>(16);

		// 标签是否处于表格内
		if (bookMark.isInTable()) {
			// 获得标签对应的Table对象和Row对象
			XWPFTable table = bookMark.getContainerTable();
			XWPFTableRow row = bookMark.getContainerTableRow();

			CTRow ctRow = row.getCtRow();
			List<XWPFTableCell> rowCell = row.getTableCells();
			for (int i = 0; i < rowCell.size(); i++) {
				columnMap.put(i + "", rowCell.get(i).getText().trim());

				// 获取该单元格段落的xml，得到根节点
				Node node1 = rowCell.get(i).getParagraphs().get(0).getCTP().getDomNode();
				// 遍历根节点的所有子节点
				for (int x = 0; x < node1.getChildNodes().getLength(); x++) {
					if (node1.getChildNodes().item(x).getNodeName()
							.equals(BookMark.RUN_NODE_NAME)) {
						Node node2 = node1.getChildNodes().item(x);

						// 遍历所有节点为"w:r"的所有自己点，找到节点名为"w:rPr"的节点
						for (int y = 0; y < node2.getChildNodes().getLength(); y++) {
							if (node2.getChildNodes().item(y).getNodeName()
									.endsWith(BookMark.STYLE_NODE_NAME)) {
								// 将节点为"w:rPr"的节点(字体格式)存到HashMap中
								styleNode.put(i + "", node2.getChildNodes().item(y));
							}
						}
					}
				}
			}

			// 循环对比，找到标签行所处的位置，rowNum为行数
			for (int i = 0; i < table.getNumberOfRows(); i++) {
				if (table.getRow(i).equals(row)) {
					rowNum = i;
					break;
				}
			}

			//循环次数为插入数据条数
			for (int i = 0; i < content.size(); i++) {
				XWPFTableRow newTableRow = table.insertNewTableRow(rowNum+i+1);
				if (newTableRow == null) {
					return;
				}

				CTTbl ctTbl = table.getCTTbl();
				CTTblGrid tblGrid = ctTbl.getTblGrid();
				if (tblGrid != null) {
					// 新增单元格
					List<CTTblGridCol> gridColList = tblGrid.getGridColList();
					if (gridColList != null && gridColList.size() > 0) {
						for (CTTblGridCol ctlCol : gridColList) {
							XWPFTableCell cell = newTableRow.addNewTableCell();
						}

						//获得插入新行的单元格数，循环插入新值，设置单元格格式
						List<XWPFTableCell> cells = newTableRow.getTableCells();
						for (int j = 0; j < cells.size(); j++) {
							XWPFParagraph para = cells.get(j).getParagraphs().get(0);
							//居中，行高
							para.setAlignment(ParagraphAlignment.CENTER);
							PoiWordUtil.setParHeight(para,"240");

							//文字在表格内居中
							CTTc cttc = cells.get(j).getCTTc();
							CTTcPr ctPr = cttc.addNewTcPr();
							ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);

							cttc.getPArray(0).addNewPPr().addNewJc().setVal(STJc.CENTER);

							if (content.get(i).get(columnMap.get(j + "")) != null) {
								XWPFRun run = para.insertNewRun(0);
								// 填入值，格式设置为标题的格式
								run.setText(content.get(i).get(columnMap.get(j + ""))+ "");
								run.getCTR().getDomNode().insertBefore(styleNode.get(j + "").cloneNode(true),
										run.getCTR().getDomNode().getFirstChild());
							}

						}
					}
				}
			}
			if("鉴定材料".equals(bookMarkName)) {
				//行合并
				poiWordUtil.mergeCellsVertically(table,0,rowNum,rowNum+content.size());
			}

		}

	}


	public void saveAs(String saveUrl,String destDirName) {
		File newFile = new File(saveUrl);
		File newFilePath = new File(destDirName);
		if (!destDirName.endsWith(File.separator)) {  
            destDirName = destDirName + File.separator;  
        }
		newFilePath.mkdirs();
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(newFile);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		try {
			this.document.write(fos);
			fos.flush();
			fos.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	/**
	 * 进行标签替换的例子,传入的Map中，key表示标签名称，value是替换的信息
	 * 对应标签插入图片
	 * @param indicator
	 */
	public void replaceBookMarkPhoto(Map<String, String> indicator, Map<String,String> map) {
		// 循环进行替换
		Iterator<String> bookMarkIter = bookMarks.getNameIterator();
		while (bookMarkIter.hasNext()) {
			String bookMarkName = bookMarkIter.next();

			// 得到标签名称
			BookMark bookMark = bookMarks.getBookmark(bookMarkName);

			// 进行替换
			if (indicator.get(bookMarkName) != null) {
				bookMark.insertPhotoAtBookMark(indicator.get(bookMarkName),map);
			}
		}
	}



}
