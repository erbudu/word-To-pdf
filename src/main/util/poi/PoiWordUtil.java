package main.util.poi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.Borders;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.VerticalAlign;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRelation;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlToken;
import org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTColor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTEm;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFldChar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHighlight;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHpsMeasure;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTInd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTMarkupRange;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTParaRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSignedTwipsMeasure;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSpacing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTString;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTextScale;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTUnderline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STEm;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STFldCharType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHdrFtr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHighlightColor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STLineSpacingRule;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STUnderline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalAlignRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

/** 报表导出样式
 * @author Se7en
 */
public class PoiWordUtil {
	/**
	 * 跨列合并
	 * 
	 * @param table
	 * @param row
	 * @param fromCell
	 */
	public void mergeCellsHorizontal(XWPFTable table, int row, int fromCell,
			int toCell) {
		for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++) {
			XWPFTableCell cell = table.getRow(row).getCell(cellIndex);
			if (cellIndex == fromCell) {
				// The first merged cell is set with RESTART merge value
				cell.getCTTc().addNewTcPr().addNewHMerge()
						.setVal(STMerge.RESTART);
			} else {
				// Cells which join (merge) the first one, are set with CONTINUE
				cell.getCTTc().addNewTcPr().addNewHMerge()
						.setVal(STMerge.CONTINUE);
			}
		}
	}

	/**
	 * 跨行合并
	 * 
	 * @param table
	 * @param col
	 * @param fromRow
	 * @param toRow
	 */
	public void mergeCellsVertically(XWPFTable table, int col, int fromRow,
			int toRow) {
		for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
			XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
			if (rowIndex == fromRow) {
				// The first merged cell is set with RESTART merge value
				cell.getCTTc().addNewTcPr().addNewVMerge()
						.setVal(STMerge.RESTART);
			} else {
				// Cells which join (merge) the first one, are set with CONTINUE
				cell.getCTTc().addNewTcPr().addNewVMerge()
						.setVal(STMerge.CONTINUE);
			}
		}
	}

	/**
	 * 设置表格宽度 居中
	 * 
	 * @param table
	 * @param width
	 */
	public void setTableWidth(XWPFTable table, String width) {
		CTTbl ttbl = table.getCTTbl();
		CTTblPr tblPr = ttbl.getTblPr() == null ? ttbl.addNewTblPr() : ttbl
				.getTblPr();
		CTTblWidth tblWidth = tblPr.isSetTblW() ? tblPr.getTblW() : tblPr
				.addNewTblW();
		CTJc cTJc = tblPr.addNewJc();
		cTJc.setVal(STJc.Enum.forString("center"));
		tblWidth.setW(new BigInteger(width));
		tblWidth.setType(STTblWidth.AUTO);
	}

	/**
	 * 设置单元格背景色
	 * 
	 * @param cell
	 * @param bgcolor
	 */
	public void setCellBgcolor(XWPFTableCell cell, String bgcolor) {
		CTTc cttc = cell.getCTTc();
		CTTcPr ctPr = cttc.addNewTcPr();
		CTShd ctshd = ctPr.addNewShd();
		ctshd.setFill(bgcolor);
	}

	/**
	 * 设置单元格字体格式
	 * 
	 * @param cell
	 * @param cellText
	 */
	public void setParagraph(XWPFTableCell cell, String cellText, String font,
			boolean bold) {
		CTP ctp = CTP.Factory.newInstance();
		XWPFParagraph p = new XWPFParagraph(ctp, cell);
		p.setAlignment(ParagraphAlignment.CENTER);
		XWPFRun run = p.createRun();
		cellText.trim();
		String[] str = cellText.split("\\[rn\\]");
		for(int i=0;i<str.length;i++){
			run.setText(str[i]);
			if(i+1<str.length){
				run.addBreak();
			}
			
		}
		run.setBold(bold);// 加粗
		CTRPr rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run
				.getCTR().addNewRPr();
		CTFonts fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr
				.addNewRFonts();
		if(font.equals("")){
			
		}else{
			fonts.setAscii(font);
			fonts.setEastAsia(font);
			fonts.setHAnsi(font);
		}
		
		cell.setParagraph(p);
	}

	/**
	 * 设置等保测评表格样式一
	 * 
	 * @param cell
	 * @param text
	 * @param bgcolor
	 * @param width
	 */
	public void setCellText(XWPFTableCell cell, String text, String bgcolor,
			int width, String font,String fontSize, boolean bold) {
		CTTc cttc = cell.getCTTc();
		CTTcPr cellPr = cttc.addNewTcPr();
		cellPr.addNewTcW().setW(BigInteger.valueOf(width));
		cell.setColor(bgcolor);
		
		CTTcPr ctPr = cttc.addNewTcPr();
		CTShd ctshd = ctPr.addNewShd();
		ctshd.setFill(bgcolor);
		ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
		cttc.getPArray(0).addNewPPr().addNewJc().setVal(STJc.CENTER);

		CTP ctp = CTP.Factory.newInstance();
		XWPFParagraph p = new XWPFParagraph(ctp, cell);
		p.setAlignment(ParagraphAlignment.CENTER);
		XWPFRun run = p.createRun();
		text.trim();
		String[] str = text.split("\\[rn\\]");
		for(int i=0;i<str.length;i++){
			run.setText(str[i]);
			if(i+1<str.length){
				run.addBreak();
			}
			
		}
		run.setBold(bold);
		
		CTRPr rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run
				.getCTR().addNewRPr();
		//根据fontSize源码方法改动
		CTHpsMeasure ctSize = rpr.isSetSz() ? rpr.getSz() : rpr.addNewSz();
		ctSize.setVal(new BigInteger(fontSize));
		
		CTFonts fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr
				.addNewRFonts();
		if(font.equals("")){
			
		}else{
			fonts.setAscii(font);
			fonts.setEastAsia(font);
			fonts.setHAnsi(font);
		}
		
		cell.setParagraph(p);
	}

	/**
	 * 设置段落对齐方式
	 * 
	 * @param p
	 * @param pAlign
	 * @param valign
	 *            poiWordUtil.setParagraphAlignInfo(p,
	 *            ParagraphAlignment.CENTER, TextAlignment.CENTER);
	 */
	public void setParagraphAlignInfo(XWPFParagraph p,
			ParagraphAlignment pAlign, TextAlignment valign) {
		p.setAlignment(pAlign);
		p.setVerticalAlignment(valign);
	}

	/**
	 * 添加页脚：显示页码信息
	 * 
	 * @param document
	 * @throws Exception
	 */
	public void simpleNumberFooter(XWPFDocument document) throws Exception {
		CTP ctp = CTP.Factory.newInstance();
		XWPFParagraph codePara = new XWPFParagraph(ctp, document);
		XWPFRun r1 = codePara.createRun();
		r1.setText("第");
		r1.setFontSize(11);
		CTRPr rpr = r1.getCTR().isSetRPr() ? r1.getCTR().getRPr() : r1.getCTR()
				.addNewRPr();
		CTFonts fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr
				.addNewRFonts();
		fonts.setAscii("宋体");
		fonts.setEastAsia("宋体");
		fonts.setHAnsi("宋体");

		r1 = codePara.createRun();
		CTFldChar fldChar = r1.getCTR().addNewFldChar();
		fldChar.setFldCharType(STFldCharType.Enum.forString("begin"));

		r1 = codePara.createRun();
		CTText ctText = r1.getCTR().addNewInstrText();
		ctText.setStringValue("PAGE  \\* MERGEFORMAT");
		ctText.setSpace(SpaceAttribute.Space.Enum.forString("preserve"));
		r1.setFontSize(11);
		rpr = r1.getCTR().isSetRPr() ? r1.getCTR().getRPr() : r1.getCTR()
				.addNewRPr();
		fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr.addNewRFonts();
		fonts.setAscii("宋体");
		fonts.setEastAsia("宋体");
		fonts.setHAnsi("宋体");

		fldChar = r1.getCTR().addNewFldChar();
		fldChar.setFldCharType(STFldCharType.Enum.forString("end"));

		r1 = codePara.createRun();
		r1.setText("页 总共");
		r1.setFontSize(11);
		rpr = r1.getCTR().isSetRPr() ? r1.getCTR().getRPr() : r1.getCTR()
				.addNewRPr();
		fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr.addNewRFonts();
		fonts.setAscii("宋体");
		fonts.setEastAsia("宋体");
		fonts.setHAnsi("宋体");

		r1 = codePara.createRun();
		fldChar = r1.getCTR().addNewFldChar();
		fldChar.setFldCharType(STFldCharType.Enum.forString("begin"));

		r1 = codePara.createRun();
		ctText = r1.getCTR().addNewInstrText();
		ctText.setStringValue("NUMPAGES  \\* MERGEFORMAT ");
		ctText.setSpace(SpaceAttribute.Space.Enum.forString("preserve"));
		r1.setFontSize(11);
		rpr = r1.getCTR().isSetRPr() ? r1.getCTR().getRPr() : r1.getCTR()
				.addNewRPr();
		fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr.addNewRFonts();
		fonts.setAscii("宋体");
		fonts.setEastAsia("宋体");
		fonts.setHAnsi("宋体");

		fldChar = r1.getCTR().addNewFldChar();
		fldChar.setFldCharType(STFldCharType.Enum.forString("end"));

		r1 = codePara.createRun();
		r1.setText("页");
		r1.setFontSize(11);
		rpr = r1.getCTR().isSetRPr() ? r1.getCTR().getRPr() : r1.getCTR()
				.addNewRPr();
		fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr.addNewRFonts();
		fonts.setAscii("宋体");
		fonts.setEastAsia("宋体");
		fonts.setHAnsi("宋体");

		codePara.setAlignment(ParagraphAlignment.CENTER);
		codePara.setVerticalAlignment(TextAlignment.CENTER);
		codePara.setBorderTop(Borders.THICK);
		XWPFParagraph[] newparagraphs = new XWPFParagraph[1];
		newparagraphs[0] = codePara;
		CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
		XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(
				document, sectPr);
		headerFooterPolicy.createFooter(STHdrFtr.DEFAULT, newparagraphs);
	}

	/**
	 * 添加时间头 相当于添加页头
	 * 
	 * @param document
	 * @throws Exception
	 */
	public void simpleDateHeader(XWPFDocument document) throws Exception {
		CTP ctp = CTP.Factory.newInstance();
		XWPFParagraph codePara = new XWPFParagraph(ctp, document);

		XWPFRun r1 = codePara.createRun();
		CTFldChar fldChar = r1.getCTR().addNewFldChar();
		fldChar.setFldCharType(STFldCharType.Enum.forString("begin"));

		r1 = codePara.createRun();
		CTText ctText = r1.getCTR().addNewInstrText();
		ctText.setStringValue("TIME \\@ \"EEEE\"");
		ctText.setSpace(SpaceAttribute.Space.Enum.forString("preserve"));
		r1.setFontSize(11);
		CTRPr rpr = r1.getCTR().isSetRPr() ? r1.getCTR().getRPr() : r1.getCTR()
				.addNewRPr();
		CTFonts fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr
				.addNewRFonts();
		fonts.setAscii("微软雅黑");
		fonts.setEastAsia("微软雅黑");
		fonts.setHAnsi("微软雅黑");

		fldChar = r1.getCTR().addNewFldChar();
		fldChar.setFldCharType(STFldCharType.Enum.forString("end"));

		r1 = codePara.createRun();
		r1.setText("年");
		r1.setFontSize(11);
		rpr = r1.getCTR().isSetRPr() ? r1.getCTR().getRPr() : r1.getCTR()
				.addNewRPr();
		fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr.addNewRFonts();
		fonts.setAscii("微软雅黑");
		fonts.setEastAsia("微软雅黑");
		fonts.setHAnsi("微软雅黑");

		r1 = codePara.createRun();
		fldChar = r1.getCTR().addNewFldChar();
		fldChar.setFldCharType(STFldCharType.Enum.forString("begin"));

		r1 = codePara.createRun();
		ctText = r1.getCTR().addNewInstrText();
		ctText.setStringValue("TIME \\@ \"O\"");
		ctText.setSpace(SpaceAttribute.Space.Enum.forString("preserve"));
		r1.setFontSize(11);
		rpr = r1.getCTR().isSetRPr() ? r1.getCTR().getRPr() : r1.getCTR()
				.addNewRPr();
		fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr.addNewRFonts();
		fonts.setAscii("微软雅黑");
		fonts.setEastAsia("微软雅黑");
		fonts.setHAnsi("微软雅黑");

		fldChar = r1.getCTR().addNewFldChar();
		fldChar.setFldCharType(STFldCharType.Enum.forString("end"));

		r1 = codePara.createRun();
		r1.setText("月");
		r1.setFontSize(11);
		rpr = r1.getCTR().isSetRPr() ? r1.getCTR().getRPr() : r1.getCTR()
				.addNewRPr();
		fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr.addNewRFonts();
		fonts.setAscii("微软雅黑");
		fonts.setEastAsia("微软雅黑");
		fonts.setHAnsi("微软雅黑");

		r1 = codePara.createRun();
		fldChar = r1.getCTR().addNewFldChar();
		fldChar.setFldCharType(STFldCharType.Enum.forString("begin"));

		r1 = codePara.createRun();
		ctText = r1.getCTR().addNewInstrText();
		ctText.setStringValue("TIME \\@ \"A\"");
		ctText.setSpace(SpaceAttribute.Space.Enum.forString("preserve"));
		r1.setFontSize(11);
		rpr = r1.getCTR().isSetRPr() ? r1.getCTR().getRPr() : r1.getCTR()
				.addNewRPr();
		fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr.addNewRFonts();
		fonts.setAscii("微软雅黑");
		fonts.setEastAsia("微软雅黑");
		fonts.setHAnsi("微软雅黑");

		fldChar = r1.getCTR().addNewFldChar();
		fldChar.setFldCharType(STFldCharType.Enum.forString("end"));

		r1 = codePara.createRun();
		r1.setText("日");
		r1.setFontSize(11);
		rpr = r1.getCTR().isSetRPr() ? r1.getCTR().getRPr() : r1.getCTR()
				.addNewRPr();
		fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr.addNewRFonts();
		fonts.setAscii("微软雅黑");
		fonts.setEastAsia("微软雅黑");
		fonts.setHAnsi("微软雅黑");

		codePara.setAlignment(ParagraphAlignment.CENTER);
		codePara.setVerticalAlignment(TextAlignment.CENTER);
		codePara.setBorderBottom(Borders.THICK);
		XWPFParagraph[] newparagraphs = new XWPFParagraph[1];
		newparagraphs[0] = codePara;
		CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
		XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(
				document, sectPr);
		headerFooterPolicy.createHeader(STHdrFtr.DEFAULT, newparagraphs);
	}

	/**
	 * 设置单元格格式及背景色
	 * 
	 * @param cell
	 * @param text
	 * @param bgcolor
	 * @param width
	 */
	public void setCellText(XWPFTableCell cell, String text, String bgcolor,int width) {
		CTTc cttc = cell.getCTTc();
		CTTcPr ctPr = cttc.addNewTcPr();
		CTShd ctshd = ctPr.addNewShd();
		ctPr.addNewTcW().setW(BigInteger.valueOf(width));
		ctshd.setFill(bgcolor);
		ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
		cttc.getPArray(0).addNewPPr().addNewJc().setVal(STJc.CENTER);
		cell.setText(text);
	}

	/**
	 * 设置段落边框
	 * 
	 * @param p
	 * @param lborder
	 * @param tBorders
	 * @param rBorders
	 * @param bBorders
	 * @param btborders
	 */
	public void setParagraphBorder(XWPFParagraph p, Borders lborder,
			Borders tBorders, Borders rBorders, Borders bBorders,
			Borders btborders) {
		if (lborder != null) {
			p.setBorderLeft(lborder);
		}
		if (tBorders != null) {
			p.setBorderTop(tBorders);
		}
		if (rBorders != null) {
			p.setBorderRight(rBorders);
		}
		if (bBorders != null) {
			p.setBorderBottom(bBorders);
		}
		if (btborders != null) {
			p.setBorderBetween(btborders);
		}
	}

	/**
	 * 设置段落缩进信息 1厘米≈567
	 * 
	 * @param p
	 * @param firstLine
	 * @param firstLineChar
	 * @param hanging
	 * @param hangingChar
	 * @param right
	 * @param rigthChar
	 * @param left
	 * @param leftChar
	 */
	public void setParagraphIndInfo(XWPFParagraph p, String firstLine,
			String firstLineChar, String hanging, String hangingChar,
			String right, String rigthChar, String left, String leftChar) {
		CTPPr pPPr = null;
		if (p.getCTP() != null) {
			if (p.getCTP().getPPr() != null) {
				pPPr = p.getCTP().getPPr();
			} else {
				pPPr = p.getCTP().addNewPPr();
			}
		}
		CTInd pInd = pPPr.getInd() != null ? pPPr.getInd() : pPPr.addNewInd();
		if (firstLine != null) {
			pInd.setFirstLine(new BigInteger(firstLine));
		}
		if (firstLineChar != null) {
			pInd.setFirstLineChars(new BigInteger(firstLineChar));
		}
		if (hanging != null) {
			pInd.setHanging(new BigInteger(hanging));
		}
		if (hangingChar != null) {
			pInd.setHangingChars(new BigInteger(hangingChar));
		}
		if (left != null) {
			pInd.setLeft(new BigInteger(left));
		}
		if (leftChar != null) {
			pInd.setLeftChars(new BigInteger(leftChar));
		}
		if (right != null) {
			pInd.setRight(new BigInteger(right));
		}
		if (rigthChar != null) {
			pInd.setRightChars(new BigInteger(rigthChar));
		}
	}

	/**
	 * 设置段落间距信息 一行=240 1.5行倍距 =360  一磅=20
	 * 
	 * @param p
	 * @param isSpace
	 * @param before
	 * @param after
	 * @param beforeLines
	 * @param afterLines
	 * @param isLine
	 * @param line
	 * @param lineValue
	 */
	public void setParagraphSpacingInfo(XWPFParagraph p, boolean isSpace,
			String before, String after, String beforeLines, String afterLines,
			boolean isLine, String line, STLineSpacingRule.Enum lineValue) {
		CTPPr pPPr = null;
		if (p.getCTP() != null) {
			if (p.getCTP().getPPr() != null) {
				pPPr = p.getCTP().getPPr();
			} else {
				pPPr = p.getCTP().addNewPPr();
			}
		}
		CTSpacing pSpacing = pPPr.getSpacing() != null ? pPPr.getSpacing()
				: pPPr.addNewSpacing();
		if (isSpace) {
			// 段前磅数
			if (before != null) {
				pSpacing.setBefore(new BigInteger(before));
			}
			// 段后磅数
			if (after != null) {
				pSpacing.setAfter(new BigInteger(after));
			}
			// 段前行数
			if (beforeLines != null) {
				pSpacing.setBeforeLines(new BigInteger(beforeLines));
			}
			// 段后行数
			if (afterLines != null) {
				pSpacing.setAfterLines(new BigInteger(afterLines));
			}
		}
		if (isLine) {
			if (line != null) {
				pSpacing.setLine(new BigInteger(line));
			}
			if (lineValue != null) {
				pSpacing.setLineRule(lineValue);
			}
		}
	}

	/**
	 * 设置字体信息 设置字符间距信息(CTSignedTwipsMeasure)
	 * 
	 * @param verticalAlign
	 *            SUPERSCRIPT上标 SUBSCRIPT下标
	 * @param position
	 *            字符位置 1磅=2
	 */
	public void setTextFontInfo(XWPFParagraph p, boolean isInsert,
			boolean isNewLine, String content, String fontFamily,
			String colorVal, String fontSize, boolean isBlod,
			UnderlinePatterns underPatterns, boolean isItalic,
			boolean isStrike, VerticalAlign verticalAlign, int position,
			String spacingValue) {
		XWPFRun pRun = null;
		if (isInsert) {
			pRun = p.createRun();
		} else {
			if (p.getRuns() != null && p.getRuns().size() > 0) {
				pRun = p.getRuns().get(0);
			} else {
				pRun = p.createRun();
			}
		}
		if (isNewLine) {
			pRun.addBreak();
		}
		pRun.setText(content);
		// 设置字体样式
		pRun.setBold(isBlod);
		pRun.setItalic(isItalic);
		if (underPatterns != null) {
			pRun.setUnderline(underPatterns);
		}
		pRun.setColor(colorVal);
		if (verticalAlign != null) {
			pRun.setSubscript(verticalAlign);
		}
		pRun.setTextPosition(position);

		CTRPr pRpr = null;
		if (pRun.getCTR() != null) {
			pRpr = pRun.getCTR().getRPr();
			if (pRpr == null) {
				pRpr = pRun.getCTR().addNewRPr();
			}
		} else {
			// pRpr = p.getCTP().addNewR().addNewRPr();
		}
		// 设置字体
		CTFonts fonts = pRpr.isSetRFonts() ? pRpr.getRFonts() : pRpr
				.addNewRFonts();
		if(fontFamily.equals("")){
			
		}else{
			fonts.setAscii(fontFamily);
			fonts.setEastAsia(fontFamily);
			fonts.setHAnsi(fontFamily);
		}
		
		// 设置字体大小
		CTHpsMeasure sz = pRpr.isSetSz() ? pRpr.getSz() : pRpr.addNewSz();
		sz.setVal(new BigInteger(fontSize));

		CTHpsMeasure szCs = pRpr.isSetSzCs() ? pRpr.getSzCs() : pRpr
				.addNewSzCs();
		szCs.setVal(new BigInteger(fontSize));

		if (spacingValue != null) {
			// 设置字符间距信息
			CTSignedTwipsMeasure ctSTwipsMeasure = pRpr.isSetSpacing() ? pRpr
					.getSpacing() : pRpr.addNewSpacing();
			ctSTwipsMeasure.setVal(new BigInteger(spacingValue));
		}
	}

	/**
	 * 为段落添加超链接
	 * 
	 * @Description: 添加超链接
	 * @param position
	 *            1磅=2
	 */
	public void appendExternalHyperlink(String url, String text,
			XWPFParagraph paragraph, String fontFamily, String fontSize,
			boolean isBlod, boolean isItalic, boolean isStrike,
			String verticalAlign, String position, String spacingValue) {
		// Add the link as External relationship
		String id = paragraph
				.getDocument()
				.getPackagePart()
				.addExternalRelationship(url,
						XWPFRelation.HYPERLINK.getRelation()).getId();
		// Append the link and bind it to the relationship
		CTHyperlink cLink = paragraph.getCTP().addNewHyperlink();
		cLink.setId(id);

		// Create the linked text
		CTText ctText = CTText.Factory.newInstance();
		ctText.setStringValue(text);
		CTR ctr = CTR.Factory.newInstance();
		CTRPr rpr = ctr.addNewRPr();

		// 设置超链接样式
		CTColor color = CTColor.Factory.newInstance();
		color.setVal("0000FF");
		rpr.setColor(color);
		rpr.addNewU().setVal(STUnderline.SINGLE);
		if (isBlod) {
			rpr.addNewB().setVal(STOnOff.Enum.forString("true"));
		}
		if (isItalic) {
			rpr.addNewI().setVal(STOnOff.Enum.forString("true"));
		}
		if (isStrike) {
			rpr.addNewStrike().setVal(STOnOff.Enum.forString("true"));
		}
		if (verticalAlign != null) {
			rpr.addNewVertAlign().setVal(
					STVerticalAlignRun.Enum.forString(verticalAlign));
		}
		rpr.addNewPosition().setVal(new BigInteger(position));

		// 设置字体
		CTFonts fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr
				.addNewRFonts();
		fonts.setAscii(fontFamily);
		fonts.setEastAsia(fontFamily);
		fonts.setHAnsi(fontFamily);

		// 设置字体大小
		CTHpsMeasure sz = rpr.isSetSz() ? rpr.getSz() : rpr.addNewSz();
		sz.setVal(new BigInteger(fontSize));

		CTHpsMeasure szCs = rpr.isSetSzCs() ? rpr.getSzCs() : rpr.addNewSzCs();
		szCs.setVal(new BigInteger(fontSize));

		if (spacingValue != null) {
			// 设置字符间距信息
			CTSignedTwipsMeasure ctSTwipsMeasure = rpr.isSetSpacing() ? rpr
					.getSpacing() : rpr.addNewSpacing();
			ctSTwipsMeasure.setVal(new BigInteger(spacingValue));
		}

		ctr.setTArray(new CTText[] { ctText });
		cLink.setRArray(new CTR[] { ctr });
	}

	/**
	 * 添加新页面
	 * 
	 * @param document
	 * @param breakType
	 *            addNewPage(document, BreakType.PAGE)
	 */
	public void addNewPage(XWPFDocument document, BreakType breakType) {
		XWPFParagraph xp = document.createParagraph();
		xp.createRun().addBreak(breakType);
	}

	/**
	 * 设置页边距 设置页边距 1厘米约等于567
	 * 
	 * @param document
	 * @param left
	 * @param top
	 * @param right
	 * @param bottom
	 *            setDocumentMargin(document, "1797", "1440", "1797", "1440");
	 */
	public void setDocumentMargin(XWPFDocument document, String left,
			String top, String right, String bottom) {
		CTSectPr sectPr = document.getDocument().getBody().isSetSectPr() ? document
				.getDocument().getBody().getSectPr()
				: document.getDocument().getBody().addNewSectPr();
		CTPageMar ctpagemar = sectPr.addNewPgMar();
		ctpagemar.setLeft(new BigInteger(left));
		ctpagemar.setTop(new BigInteger(top));
		ctpagemar.setRight(new BigInteger(right));
		ctpagemar.setBottom(new BigInteger(bottom));
	}

	/**
	 * 保存文档
	 * 
	 * @param document
	 * @param savePath
	 * @throws Exception
	 */
	public void saveDocument(XWPFDocument document, String savePath)
			throws Exception {
		FileOutputStream fos = new FileOutputStream(savePath);
		document.write(fos);
		fos.close();
	}

	/**
	 * 网安多级编号样式
	 * 
	 * @param doc
	 * @param Level
	 *            0一级 1二级
	 * @param MutiLevelText
	 */
	public void createMutiLevel(XWPFDocument doc, int Level,
			String MutiLevelText) {
		XWPFParagraph para2 = doc.createParagraph();
		para2.setStyle("2");
//		para2.setNumID(BigInteger.valueOf(2));
//		para2.getCTP().getPPr().getNumPr().addNewIlvl()
//				.setVal(BigInteger.valueOf(Level));// 设置等级
		XWPFRun run2 = para2.createRun();
		run2.setText(MutiLevelText);
		run2.setBold(true);
		run2.setFontFamily("黑体");

	}

	/**
	 * 网安多级编号样式
	 * 
	 * @param doc
	 * @param styleID
	 * @param mutiLevelText
	 */
	public void createMutiLevel(XWPFDocument doc,String styleID, String mutiLevelText,String color,String font,int fontsize) {
		XWPFParagraph para2 = doc.createParagraph();
		para2.setStyle(styleID);
		XWPFRun run2 = para2.createRun();
		run2.setText(mutiLevelText);
		run2.setBold(true);
		
		// 设置字体大小
				
		if(font.equals("")){
		}else{
			run2.setFontFamily(font);
		}
		
		if(color.equals("")){
		}else{
			run2.setColor(color);
		}
			run2.setFontSize(fontsize);
	}

	/** 
     * 替换段落里面的变量 
     * 
     * @param doc    要替换的文档 
     * @param params 参数 
     */  
    public void replaceInPara(XWPFDocument doc, Map<String, Object> params) {  
        Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();  
        XWPFParagraph para;  
        while (iterator.hasNext()) {  
            para = iterator.next();  
            this.replaceInPara(para, params);  
        }  
    }  
  
    /** 
     * 替换段落里面的变量 
     * 
     * @param para   要替换的段落 
     * @param params 参数 
     */  
    public void replaceInPara(XWPFParagraph para, Map<String, Object> params) {  
        List<XWPFRun> runs;  
        Matcher matcher;  
        if (this.matcher(para.getParagraphText()).find()) {  
            runs = para.getRuns();  
  
            int start = -1;  
            int end = -1;  
            String str = "";  
            for (int i = 0; i < runs.size(); i++) {  
                XWPFRun run = runs.get(i);  
                String runText = run.toString();  
                if ('$' == runText.charAt(0)&&'{' == runText.charAt(1)) {  
                    start = i;  
                }  
                if ((start != -1)) {  
                    str += runText;  
                }  
                if ('}' == runText.charAt(runText.length() - 1)) {  
                    if (start != -1) {  
                        end = i;  
                        break;  
                    }  
                }  
            }  
            for (int i = start; i <= end; i++) {  
                para.removeRun(i);  
                i--;  
                end--;  
            }  
  
            for (String key : params.keySet()) {  
                if (str.equals(key)) {  
                    para.createRun().setText((String) params.get(key));  
                    break;  
                }  
            }  
  
  
        }  
    }  
  
    /** 
     * 替换表格里面的变量 
     * 
     * @param doc    要替换的文档 
     * @param params 参数 
     */  
    public void replaceInTable(XWPFDocument doc, Map<String, Object> params) {  
        Iterator<XWPFTable> iterator = doc.getTablesIterator();  
        XWPFTable table;  
        List<XWPFTableRow> rows;  
        List<XWPFTableCell> cells;  
        List<XWPFParagraph> paras;  
        while (iterator.hasNext()) {  
            table = iterator.next();  
            rows = table.getRows();  
            for (XWPFTableRow row : rows) {  
                cells = row.getTableCells();  
                for (XWPFTableCell cell : cells) {  
                    paras = cell.getParagraphs();  
                    for (XWPFParagraph para : paras) {  
                        this.replaceInPara(para, params);  
                    }  
                }  
            }  
        }  
    }  
  
    /** 
     * 正则匹配字符串 
     * 
     * @param str 
     * @return 
     */  
    private Matcher matcher(String str) {  
        Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);  
        Matcher matcher = pattern.matcher(str);  
        return matcher;  
    }  
    
    public void setTableBorders(XWPFTable table, CTBorder left, CTBorder top,
			CTBorder right, CTBorder bottom) {
		CTTblBorders tblBorders = getTableBorders(table);
		if (left != null) {
			tblBorders.setLeft(left);
		}
		if (top != null) {
			tblBorders.setTop(top);
		}
		if (right != null) {
			tblBorders.setRight(right);
		}
		if (bottom != null) {
			tblBorders.setBottom(bottom);
		}
	}
    /**
	 * @Description: 得到Table的边框,不存在则新建
	 */
	public CTTblBorders getTableBorders(XWPFTable table) {
		CTTblPr tblPr = getTableCTTblPr(table);
		CTTblBorders tblBorders = tblPr.isSetTblBorders() ? tblPr
				.getTblBorders() : tblPr.addNewTblBorders();
		return tblBorders;
	}
	
	/**
	 * @Description: 得到Table的CTTblPr,不存在则新建
	 */
	public CTTblPr getTableCTTblPr(XWPFTable table) {
		CTTbl ttbl = table.getCTTbl();
		CTTblPr tblPr = ttbl.getTblPr() == null ? ttbl.addNewTblPr() : ttbl
				.getTblPr();
		return tblPr;
	}
  
	/**
	 * word插入图片
	 * 
	 * @throws Exception
	 */
	public void insertImg(XWPFDocument doc,List<String> picUrl) throws Exception {
		XWPFParagraph p = doc.createParagraph();
		XWPFRun r = p.createRun();
		
		for (String imgFile : picUrl) {
			int format;

			if (imgFile.endsWith(".emf")) {
				format = XWPFDocument.PICTURE_TYPE_EMF;
			} else if (imgFile.endsWith(".wmf")) {
				format = XWPFDocument.PICTURE_TYPE_WMF;
			} else if (imgFile.endsWith(".pict")) {
				format = XWPFDocument.PICTURE_TYPE_PICT;
			} else if (imgFile.endsWith(".jpeg") || imgFile.endsWith(".jpg")) {
				format = XWPFDocument.PICTURE_TYPE_JPEG;
			} else if (imgFile.endsWith(".png")) {
				format = XWPFDocument.PICTURE_TYPE_PNG;
			} else if (imgFile.endsWith(".dib")) {
				format = XWPFDocument.PICTURE_TYPE_DIB;
			} else if (imgFile.endsWith(".gif")) {
				format = XWPFDocument.PICTURE_TYPE_GIF;
			} else if (imgFile.endsWith(".tiff")) {
				format = XWPFDocument.PICTURE_TYPE_TIFF;
			} else if (imgFile.endsWith(".eps")) {
				format = XWPFDocument.PICTURE_TYPE_EPS;
			} else if (imgFile.endsWith(".bmp")) {
				format = XWPFDocument.PICTURE_TYPE_BMP;
			} else if (imgFile.endsWith(".wpg")) {
				format = XWPFDocument.PICTURE_TYPE_WPG;
			} else {
				System.err.println("Unsupported picture: "+ imgFile
								+ ". Expected emf|wmf|pict|jpeg|png|dib|gif|tiff|eps|bmp|wpg");
				continue;
			}
			// 200x200 pixels
			r.addPicture(new FileInputStream(imgFile), format, imgFile,
					Units.toEMU(410), Units.toEMU(250));
			r.addBreak(BreakType.PAGE);
		}


	}

	
	/**
	 * @Description: 添加书签
	 */
	public void addParagraphContentBookmark(XWPFParagraph p, String content,
			BigInteger markId, String bookMarkName, boolean isInsert,
			boolean isNewLine, String fontFamily, String fontSize,
			String colorVal, boolean isBlod, boolean isUnderLine,
			String underLineColor, STUnderline.Enum underStyle,
			boolean isItalic, boolean isStrike, boolean isDStrike,
			boolean isShadow, boolean isVanish, boolean isEmboss,
			boolean isImprint, boolean isOutline, boolean isEm,
			STEm.Enum emType, boolean isHightLight,
			STHighlightColor.Enum hightStyle, boolean isShd,
			STShd.Enum shdStyle, String shdColor, VerticalAlign verticalAlign,
			int position, int spacingValue, int indent) {
		CTBookmark bookStart = p.getCTP().addNewBookmarkStart();
		bookStart.setId(markId);
		bookStart.setName(bookMarkName);

		XWPFRun pRun = getOrAddParagraphFirstRun(p, isInsert, isNewLine);
		setParagraphRunFontInfo(p, pRun, content, fontFamily, fontSize);
		setParagraphTextStyleInfo(p, pRun, colorVal, isBlod, isUnderLine,
				underLineColor, underStyle, isItalic, isStrike, isDStrike,
				isShadow, isVanish, isEmboss, isImprint, isOutline, isEm,
				emType, isHightLight, hightStyle, isShd, shdStyle, shdColor,
				verticalAlign, position, spacingValue, indent);

		CTMarkupRange bookEnd = p.getCTP().addNewBookmarkEnd();
		bookEnd.setId(markId);

	}
	
	public XWPFRun getOrAddParagraphFirstRun(XWPFParagraph p, boolean isInsert,
			boolean isNewLine) {
		XWPFRun pRun = null;
		if (isInsert) {
			pRun = p.createRun();
		} else {
			if (p.getRuns() != null && p.getRuns().size() > 0) {
				pRun = p.getRuns().get(0);
			} else {
				pRun = p.createRun();
			}
		}
		if (isNewLine) {
			pRun.addBreak();
		}
		return pRun;
	}
	
	/**
	 * @Description 设置字体信息
	 */
	public void setParagraphRunFontInfo(XWPFParagraph p, XWPFRun pRun,
			String content, String fontFamily, String fontSize) {
		CTRPr pRpr = getRunCTRPr(p, pRun);
		if (StringUtils.isNotBlank(content)) {
			pRun.setText(content);
		}
		// 设置字体
		CTFonts fonts = pRpr.isSetRFonts() ? pRpr.getRFonts() : pRpr
				.addNewRFonts();
		fonts.setAscii(fontFamily);
		fonts.setEastAsia(fontFamily);
		fonts.setHAnsi(fontFamily);

		// 设置字体大小
		CTHpsMeasure sz = pRpr.isSetSz() ? pRpr.getSz() : pRpr.addNewSz();
		sz.setVal(new BigInteger(fontSize));

		CTHpsMeasure szCs = pRpr.isSetSzCs() ? pRpr.getSzCs() : pRpr
				.addNewSzCs();
		szCs.setVal(new BigInteger(fontSize));
	}
	
	/**
	 * @Description: 得到XWPFRun的CTRPr
	 */
	public CTRPr getRunCTRPr(XWPFParagraph p, XWPFRun pRun) {
		CTRPr pRpr = null;
		if (pRun.getCTR() != null) {
			pRpr = pRun.getCTR().getRPr();
			if (pRpr == null) {
				pRpr = pRun.getCTR().addNewRPr();
			}
		} else {
			pRpr = p.getCTP().addNewR().addNewRPr();
		}
		return pRpr;
	}
	
	
	/**
	 * @Description: 设置段落文本样式(高亮与底纹显示效果不同)设置字符间距信息(CTSignedTwipsMeasure)
	 * @param verticalAlign
	 *            : SUPERSCRIPT上标 SUBSCRIPT下标
	 * @param position
	 *            :字符间距位置：>0提升 <0降低=磅值*2 如3磅=6
	 * @param spacingValue
	 *            :字符间距间距 >0加宽 <0紧缩 =磅值*20 如2磅=40
	 * @param indent
	 *            :字符间距缩进 <100 缩
	 */
	public void setParagraphTextStyleInfo(XWPFParagraph p, XWPFRun pRun,
			String colorVal, boolean isBlod, boolean isUnderLine,
			String underLineColor, STUnderline.Enum underStyle,
			boolean isItalic, boolean isStrike, boolean isDStrike,
			boolean isShadow, boolean isVanish, boolean isEmboss,
			boolean isImprint, boolean isOutline, boolean isEm,
			STEm.Enum emType, boolean isHightLight,
			STHighlightColor.Enum hightStyle, boolean isShd,
			STShd.Enum shdStyle, String shdColor, VerticalAlign verticalAlign,
			int position, int spacingValue, int indent) {
		if (pRun == null) {
			return;
		}
		CTRPr pRpr = getRunCTRPr(p, pRun);
		if (colorVal != null) {
			pRun.setColor(colorVal);
		}
		// 设置字体样式
		// 加粗
		if (isBlod) {
			pRun.setBold(isBlod);
		}
		// 倾斜
		if (isItalic) {
			pRun.setItalic(isItalic);
		}
		// 删除线
		if (isStrike) {
			pRun.setStrike(isStrike);
		}
		// 双删除线
		if (isDStrike) {
			CTOnOff dsCtOnOff = pRpr.isSetDstrike() ? pRpr.getDstrike() : pRpr
					.addNewDstrike();
			dsCtOnOff.setVal(STOnOff.TRUE);
		}
		// 阴影
		if (isShadow) {
			CTOnOff shadowCtOnOff = pRpr.isSetShadow() ? pRpr.getShadow()
					: pRpr.addNewShadow();
			shadowCtOnOff.setVal(STOnOff.TRUE);
		}
		// 隐藏
		if (isVanish) {
			CTOnOff vanishCtOnOff = pRpr.isSetVanish() ? pRpr.getVanish()
					: pRpr.addNewVanish();
			vanishCtOnOff.setVal(STOnOff.TRUE);
		}
		// 阳文
		if (isEmboss) {
			CTOnOff embossCtOnOff = pRpr.isSetEmboss() ? pRpr.getEmboss()
					: pRpr.addNewEmboss();
			embossCtOnOff.setVal(STOnOff.TRUE);
		}
		// 阴文
		if (isImprint) {
			CTOnOff isImprintCtOnOff = pRpr.isSetImprint() ? pRpr.getImprint()
					: pRpr.addNewImprint();
			isImprintCtOnOff.setVal(STOnOff.TRUE);
		}
		// 空心
		if (isOutline) {
			CTOnOff isOutlineCtOnOff = pRpr.isSetOutline() ? pRpr.getOutline()
					: pRpr.addNewOutline();
			isOutlineCtOnOff.setVal(STOnOff.TRUE);
		}
		// 着重号
		if (isEm) {
			CTEm em = pRpr.isSetEm() ? pRpr.getEm() : pRpr.addNewEm();
			em.setVal(emType);
		}
		// 设置下划线样式
		if (isUnderLine) {
			CTUnderline u = pRpr.isSetU() ? pRpr.getU() : pRpr.addNewU();
			if (underStyle != null) {
				u.setVal(underStyle);
			}
			if (underLineColor != null) {
				u.setColor(underLineColor);
			}
		}
		// 设置突出显示文本
		if (isHightLight) {
			if (hightStyle != null) {
				CTHighlight hightLight = pRpr.isSetHighlight() ? pRpr
						.getHighlight() : pRpr.addNewHighlight();
				hightLight.setVal(hightStyle);
			}
		}
		if (isShd) {
			// 设置底纹
			CTShd shd = pRpr.isSetShd() ? pRpr.getShd() : pRpr.addNewShd();
			if (shdStyle != null) {
				shd.setVal(shdStyle);
			}
			if (shdColor != null) {
				shd.setColor(shdColor);
			}
		}
		// 上标下标
		if (verticalAlign != null) {
			pRun.setSubscript(verticalAlign);
		}
		// 设置文本位置
		pRun.setTextPosition(position);
		if (spacingValue > 0) {
			// 设置字符间距信息
			CTSignedTwipsMeasure ctSTwipsMeasure = pRpr.isSetSpacing() ? pRpr
					.getSpacing() : pRpr.addNewSpacing();
			ctSTwipsMeasure
					.setVal(new BigInteger(String.valueOf(spacingValue)));
		}
		if (indent > 0) {
			CTTextScale paramCTTextScale = pRpr.isSetW() ? pRpr.getW() : pRpr
					.addNewW();
			paramCTTextScale.setVal(indent);
		}
	}

	
	/**
	 * @Description: 添加书签
	 */
	public void addParagraphContentBookmarkBasicStyle(XWPFParagraph p,
			String content, BigInteger markId, String bookMarkName,
			boolean isInsert, boolean isNewLine, String fontFamily,
			String fontSize, String colorVal, boolean isBlod,
			boolean isUnderLine, String underLineColor,
			STUnderline.Enum underStyle, boolean isItalic, boolean isStrike) {
		CTBookmark bookStart = p.getCTP().addNewBookmarkStart();
		bookStart.setId(markId);
		bookStart.setName(bookMarkName);

		XWPFRun pRun = getOrAddParagraphFirstRun(p, isInsert, isNewLine);
		setParagraphRunFontInfo(p, pRun, content, fontFamily, fontSize);
		setParagraphTextStyleInfo(p, pRun, colorVal, isBlod, isUnderLine,
				underLineColor, underStyle, isItalic, isStrike, false, false,
				false, false, false, false, false, null, false, null, false,
				null, null, null, 0, 0, 0);
		pRun.addBreak();
		CTMarkupRange bookEnd = p.getCTP().addNewBookmarkEnd();
		bookEnd.setId(markId);
	}
	
	/**
	 * 创建页眉
	 * @param docx
	 * @param text
	 * @throws Exception
	 */
	public static void createDefaultHeader(final XWPFDocument docx, final String text) throws Exception{
	    CTP ctp = CTP.Factory.newInstance();
	    XWPFParagraph paragraph = new XWPFParagraph(ctp, docx);
	    ctp.addNewR().addNewT().setStringValue(text);
	    ctp.addNewR().addNewT().setSpace(SpaceAttribute.Space.PRESERVE);
	    CTSectPr sectPr = docx.getDocument().getBody().isSetSectPr() ? docx.getDocument().getBody().getSectPr() : docx.getDocument().getBody().addNewSectPr();
	    XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(docx, sectPr);
	    XWPFHeader header = policy.createHeader(STHdrFtr.DEFAULT, new XWPFParagraph[] { paragraph });
	    header.setXWPFDocument(docx);
	}
	/**
	 * 设置段落行高  一倍行高是240  1.5倍行高是360
	 * @param line
	 */
	public static void setParHeight(XWPFParagraph p,String line) {
		CTPPr pPPr = p.getCTP().getPPr();
		CTSpacing pSpacing = pPPr.getSpacing() != null ? pPPr.getSpacing()
				: pPPr.addNewSpacing();
		pSpacing.setLine(new BigInteger(line));
	}
	
}