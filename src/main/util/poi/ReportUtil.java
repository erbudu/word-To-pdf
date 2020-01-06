package main.util.poi;

import java.math.BigInteger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import main.util.poi.PoiWordUtil;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHeight;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTInd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTNumPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTParaRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSimpleField;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTString;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHighlightColor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHint;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STLineSpacingRule;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STStyleType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

/**
 * 该类主要是关于生成报告用到的样式的方法集合
 * @author masai
 *
 */
public class ReportUtil {
	PoiWordUtil poiWordUtil = new PoiWordUtil();

	/** 标签所在的表cell对象 **/
	private XWPFTableCell _tableCell = null;

	/**
	 * 设置单元测试表格内单个表格的样式
	 */
	public void unitTestCellFont(XWPFTableCell cell,String content,int width,String bgColor,boolean bold,ParagraphAlignment paragraphAlignment,STJc.Enum stjc) {
		CTTc cttc = cell.getCTTc();
		CTTcPr ctPr = cell.getCTTc().addNewTcPr();
		CTShd ctshd = cell.getCTTc().addNewTcPr().addNewShd();
		
		XWPFParagraph par = new XWPFParagraph(CTP.Factory.newInstance(), cell);
		XWPFRun run = par.createRun();
		run.setText(content);
		run.setBold(bold);
		par.addRun(run);
		par.setAlignment(paragraphAlignment);
		ctshd.setFill(bgColor);
		cell.setParagraph(par);
		if(width!=0){
			ctPr.addNewTcW().setW(BigInteger.valueOf(width));
		}
		ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
		cttc.getPArray(0).addNewPPr().addNewJc().setVal(stjc);
		
	}
	
	public void unitTestCellFont(XWPFTableCell cell,String content,int width,String bgColor,boolean bold,ParagraphAlignment paragraphAlignment,STJc.Enum stjc,String family,int fontSize) {
//		CTP ctp = CTP.Factory.newInstance();
		CTTc cttc = cell.getCTTc();
		CTTcPr ctPr = cell.getCTTc().addNewTcPr();
		CTShd ctshd = cell.getCTTc().addNewTcPr().addNewShd();
		
		XWPFParagraph par = new XWPFParagraph(CTP.Factory.newInstance(), cell);
		XWPFRun run = par.createRun();
		run.setText(content);
		run.setFontFamily(family);
		run.setFontSize(fontSize);
		run.setBold(bold);
		par.addRun(run);
		par.setAlignment(paragraphAlignment);
		ctshd.setFill(bgColor);
		cell.setParagraph(par);
		if(width!=0){
			ctPr.addNewTcW().setW(BigInteger.valueOf(width));
		}
		ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
		cttc.getPArray(0).addNewPPr().addNewJc().setVal(stjc);
	}
	
	/**
	 * poi 设置表格行在各页顶端以标题行形式重复出现
	 */
	public void setHeader(XWPFTableRow row,String height) {
        CTTrPr trPr = row.getCtRow().isSetTrPr() ? row.getCtRow().getTrPr() : row.getCtRow().addNewTrPr();  
        trPr.addNewTblHeader();
        if(height!=null){
        	 CTHeight h = trPr.sizeOfTrHeightArray() == 0 ? trPr.addNewTrHeight() : trPr.getTrHeightArray(0);
             h.setVal(new BigInteger(height));
        }
       
	}
	
	/** * 增加自定义标题样式。这里用的是stackoverflow的源码
	 * 
	 * @param docxDocument 目标文档
	 * @param strStyleId 样式名称
	 * @param headingLevel 样式级别
	 */
	public void addCustomHeadingStyle(XWPFDocument docxDocument, String strStyleId, int headingLevel) {

		CTStyle ctStyle = CTStyle.Factory.newInstance();
		ctStyle.setStyleId(strStyleId);

		CTString styleName = CTString.Factory.newInstance();
		styleName.setVal(strStyleId);
		ctStyle.setName(styleName);

		CTDecimalNumber indentNumber = CTDecimalNumber.Factory.newInstance();
		indentNumber.setVal(BigInteger.valueOf(headingLevel)); // lower number > style is more prominent in the formats bar
		ctStyle.setUiPriority(indentNumber);

		CTOnOff onoffnull = CTOnOff.Factory.newInstance();
		ctStyle.setUnhideWhenUsed(onoffnull); // style shows up in the formats bar
		ctStyle.setQFormat(onoffnull); // style defines a heading of the given level
		CTPPr ppr = CTPPr.Factory.newInstance();
		ppr.setOutlineLvl(indentNumber);
		ctStyle.setPPr(ppr);

		XWPFStyle style = new XWPFStyle(ctStyle); // is a null op if already defined
		XWPFStyles styles = docxDocument.createStyles();

		style.setType(STStyleType.PARAGRAPH);
		styles.addStyle(style);
	}

	/**
	 * 鉴定材料信息表表头
	 */
	public XWPFTable createMaterialHeader(XWPFParagraph par,String rsName) {
		XWPFTable table = this._tableCell.insertNewTbl(par.getCTP().newCursor());
		table.getCTTbl().getTblPr().addNewTblLayout().setType(STTblLayoutType.FIXED);
		table.setCellMargins(50, 25, 50, 25);

		CTTblBorders borders=table.getCTTbl().getTblPr().addNewTblBorders();
		//表格最上边一条线的样式
		CTBorder tBorder=borders.addNewTop();
		tBorder.setVal(STBorder.Enum.forString("NONE"));
		tBorder.setSz(new BigInteger("0"));
		tBorder.setColor("auto");
		//表格最下边一条线的样式
		CTBorder bBorder=borders.addNewBottom();
		bBorder.setVal(STBorder.Enum.forString("NONE"));
		bBorder.setSz(new BigInteger("0"));
		bBorder.setColor("auto");
		//表格最左边一条线的样式
		CTBorder lBorder=borders.addNewLeft();
		lBorder.setVal(STBorder.Enum.forString("NONE"));
		lBorder.setSz(new BigInteger("0"));
		lBorder.setColor("auto");
		//表格最右边一条线的样式
		CTBorder rBorder=borders.addNewRight();
		rBorder.setVal(STBorder.Enum.forString("NONE"));
		rBorder.setSz(new BigInteger("0"));
		rBorder.setColor("auto");
		//表格内部横向表格颜色
		CTBorder hBorder=borders.addNewInsideH();
		hBorder.setVal(STBorder.Enum.forString("single"));
		hBorder.setSz(new BigInteger("4"));
		hBorder.setColor("000000");
		//表格内部纵向表格颜色
		CTBorder vBorder=borders.addNewInsideV();
		vBorder.setVal(STBorder.Enum.forString("single"));
		vBorder.setSz(new BigInteger("4"));
		vBorder.setColor("000000");

		if(("鉴定材料").equals(rsName)) {
			poiWordUtil.setTableWidth(table, "8200");
			table.getRow(0).addNewTableCell();
			table.getRow(0).addNewTableCell();
			table.getRow(0).addNewTableCell();
			table.getRow(0).addNewTableCell();
			table.getRow(0).addNewTableCell();
			table.getRow(0).addNewTableCell();
			unitTestCellFont(table.getRow(0).getCell(0),"序号",8200/20,"FFFFFF",false,ParagraphAlignment.CENTER,STJc.CENTER);
			unitTestCellFont(table.getRow(0).getCell(1),"名称",8200/20*4,"FFFFFF",false,ParagraphAlignment.CENTER,STJc.CENTER);
			unitTestCellFont(table.getRow(0).getCell(2),"数量",8200/20,"FFFFFF",false,ParagraphAlignment.CENTER,STJc.CENTER);
			unitTestCellFont(table.getRow(0).getCell(3),"封存",8200/20*2,"FFFFFF",false,ParagraphAlignment.CENTER,STJc.CENTER);
			unitTestCellFont(table.getRow(0).getCell(4),"序列号/IMEI码",8200/20*6,"FFFFFF",false,ParagraphAlignment.CENTER,STJc.CENTER);
			unitTestCellFont(table.getRow(0).getCell(5),"备注",8200/20*3,"FFFFFF",false,ParagraphAlignment.CENTER,STJc.CENTER);
			unitTestCellFont(table.getRow(0).getCell(6),"送检日期",8200/20*3,"FFFFFF",false,ParagraphAlignment.CENTER,STJc.CENTER);
		}
		return table;
	}

	/**
	 * 传一个字符串数组  自动分行
	 * @param str
	 * @param p
	 */
	public void splitPar(String[] str,XWPFParagraph p) {
		for(String string : str){
			if(string.length()>0){
				XmlCursor cursor = p.getCTP().newCursor();
				XWPFParagraph par = p.getDocument().insertNewParagraph(cursor);
				XWPFRun run0  = par.createRun();
				run0.addTab();
				XWPFRun run1  = par.createRun();
				run1.getCTR().addNewRPr().addNewHighlight().setVal(STHighlightColor.Enum.forString("yellow"));
				run1.setText(string.trim());
				run1.setFontSize(12);
				poiWordUtil.setParagraphSpacingInfo(par, true, null, "120", null, null, true, "360", STLineSpacingRule.AUTO);
			}
		}
	}
	
	/**
	 * 自动分行  并设置样式  问题处置建议
	 * @param str
	 * @param p
	 */
	public void splitPar_iss(String[] str,XWPFParagraph p) {
		CTPPr ppr = null;
		CTParaRPr rpr = null;
		CTInd  cTInd = null;
		CTNumPr cTNumPr  = null;
		for(String string : str){
			XmlCursor cursor = p.getCTP().newCursor();
			// ---这个是关键
			XWPFParagraph par = p.getDocument().insertNewParagraph(cursor);
			XWPFRun run1  = par.createRun();
			
			ppr = par.getCTP().addNewPPr();
			ppr.addNewPStyle().setVal("a7");
			cTNumPr = ppr.addNewNumPr();
			cTNumPr.addNewIlvl().setVal(BigInteger.valueOf(0));
			cTNumPr.addNewNumId().setVal(BigInteger.valueOf(23));
			cTInd= ppr.addNewInd();
			cTInd.setLeft(BigInteger.valueOf(0));
			cTInd.setFirstLine(BigInteger.valueOf(480));
			rpr = par.getCTP().addNewPPr().addNewRPr();
			rpr.addNewRFonts().setHint(STHint.EAST_ASIA);
			run1.setText(string.trim());
			run1.setFontSize(12);
		}
	}
	
	/**
	 * 将StringBuilder 的最后一个字符替换为。
	 * @param sb
	 * @return
	 */
	public void replaceFinalString(StringBuilder sb) {
		
		Pattern patPunc = Pattern.compile("[`~!@#$^&*=|{}':;',\\[\\].<>/?~！@#￥……&*——|{}【】‘；：“'。，、？]$");
		String endStr = sb.substring(sb.length()-1, sb.length());
		Matcher matcher = patPunc.matcher(endStr);  
		if(matcher.find()){
			sb.replace(sb.length()-1, sb.length(), "。");
		}
		if(!sb.substring(sb.length()-1, sb.length()).equals("。")){
			sb.append("。");
		}
	}
	public String replaceFinalString(String str) {
		Pattern patPunc = Pattern.compile("[`~!@#$^&*=|{}':;',\\[\\].<>/?~！@#￥……&*——|{}【】‘；：”“'。，、？]$");
		String endStr = str.substring(str.length()-1, str.length());
		Matcher matcher = patPunc.matcher(endStr);  
		if(matcher.find()){
			str = str.substring(0, str.length()-1)+"。";
		}
		if(!str.substring(str.length()-1, str.length()).equals("。")){
			str=str+"。";
		}
		return str;
	}
	/**
	 * 插入表头文字
	 * @param p
	 * @param str
	 */
	public void insertTableTitle(XWPFParagraph p,String str) {
		PoiWordUtil poiWordUtil = new PoiWordUtil();
		XWPFRun run;
		CTSimpleField seq;
		run = p.createRun();
		run.setText("表格 ");
		run.setFontFamily("黑体");
		run.setFontSize(10);
		seq = p.getCTP().addNewFldSimple();
		seq.setInstr("STYLEREF 1 \\s");
		run = p.createRun();
		//此处为--
		run.setText("-");
		run.setFontFamily("黑体");
		run.setFontSize(10);
		seq = p.getCTP().addNewFldSimple();
		seq.setInstr("SEQ 表格 \\* ARABIC \\s 1");
		run = p.createRun();
		run.setText(str);
		run.setFontFamily("黑体");
		run.setFontSize(10);

        poiWordUtil.setParagraphSpacingInfo(p, true, null, "120", null, null, false, "", STLineSpacingRule.AUTO);
	}


	/**
	 * 插入图索引
	 * @param p
	 * @param str
	 */
	public void insertPhotoTitle(XWPFParagraph p,String str) {
		PoiWordUtil poiWordUtil = new PoiWordUtil();
		XWPFRun run;
		CTSimpleField seq;
		run = p.createRun();
		run.setText("图 ");
		run.setFontFamily("黑体");
		run.setFontSize(10);
		seq = p.getCTP().addNewFldSimple();
		seq.setInstr("STYLEREF 1 \\s");
		run = p.createRun();
		//此处为--
		run.setText("-");
		run.setFontFamily("黑体");
		run.setFontSize(10);
		seq = p.getCTP().addNewFldSimple();
		seq.setInstr("SEQ 图 \\* ARABIC \\s 1");
		run = p.createRun();
		run.setText(str);
		run.setFontFamily("黑体");
		run.setFontSize(10);

        poiWordUtil.setParagraphSpacingInfo(p, true, null, "120", null, null, false, "", STLineSpacingRule.AUTO);
	}

}
