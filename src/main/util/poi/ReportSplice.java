package main.util.poi;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlToken;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;

import java.io.*;
import java.util.List;

/**
 * 报告拼接
 *
 * @author Masai
 * @date 2019-03-28
 */

public class ReportSplice {

    /** 在文档中间插入另一份文档部分页面
     * 整体插入，包括正文，表格，图片，格式样式
     * 文档格式需要预先调整到一致
     * @param sourcePath 源文档路径
     * @param outPath 目标文档路径
     * @throws Exception
     */
    public static void spliceReport(String sourcePath,String outPath) throws Exception {
        InputStream srcFile = new FileInputStream(sourcePath);
        XWPFDocument docSource = new XWPFDocument(srcFile);


        InputStream targetFile = new FileInputStream(outPath);
        XWPFDocument docTarget = new XWPFDocument(targetFile);

        boolean bStart = false,bEnd = false,bTitle = false;
        String strText;

        XmlCursor cursor = null;

        for (IBodyElement bodyElement : docTarget.getBodyElements()) {
            BodyElementType elementType = bodyElement.getElementType();
            if (elementType == BodyElementType.PARAGRAPH) {
                XWPFParagraph pr = (XWPFParagraph) bodyElement;

                strText = pr.getText();
                if("[附录A完]".equals(strText)) {
                    cursor = pr.getCTP().newCursor();
                    break;
                }
            }
        }

        //获取源文档插入部分内容
        List<IBodyElement> elements = docSource.getBodyElements();
        for (int i = elements.size()-1; i>=0; i--) {
            IBodyElement bodyElement = elements.get(i);
            BodyElementType elementType = bodyElement.getElementType();
            if (elementType == BodyElementType.PARAGRAPH) {
                XWPFParagraph srcPr = (XWPFParagraph) bodyElement;
                strText = srcPr.getText();

                if(strText.contains("漏洞扫描测试")) {
                    bEnd = true;
                }

                if(strText.equals("深圳市网安计算机安全检测技术有限公司")) {
                    bStart = true;
                }

                if(!bStart) {
                    continue;
                }
                if(bEnd) {
                    if(bTitle) {
                        break;
                    } else {
                        bTitle = true;
                    }
                }

                String strSty = srcPr.getStyleID();
                if(strSty != null) {
                    copyStyle(docSource, docTarget, docSource.getStyles().getStyle(strSty));
                }
                boolean hasImage = false;

                XWPFParagraph targetPr = docTarget.insertNewParagraph(cursor);
                if(targetPr == null) {
                    break;
                }

                for (XWPFRun srcRun : srcPr.getRuns()) {
                    // You need next code when you want to call XWPFParagraph.removeRun().

                    if (srcRun.getEmbeddedPictures().size() > 0) {
                        hasImage = true;
                    }

                    for (XWPFPicture pic : srcRun.getEmbeddedPictures()) {
                        byte[] img = pic.getPictureData().getData();
                        long cx = pic.getCTPicture().getSpPr().getXfrm().getExt().getCx();
                        long cy = pic.getCTPicture().getSpPr().getXfrm().getExt().getCy();
                        XWPFRun dstRun = targetPr.createRun();
                        CTInline ctinline=dstRun.getCTR().addNewDrawing().addNewInline();

                        try {
                            // Working addPicture Code below...
                            String blipId = targetPr.getDocument().addPictureData(new ByteArrayInputStream(img),XWPFDocument.PICTURE_TYPE_PNG);
                            createPic(blipId, docTarget.getNextPicNameNumber(XWPFDocument.PICTURE_TYPE_PNG), cx, cy, ctinline);

                        } catch (InvalidFormatException e1) {
                            e1.printStackTrace();
                        }
                    }
                }

                if (!hasImage) {

                    int pos = docTarget.getParagraphs().size() - 1;
                    int npar =  docTarget.getParagraphPos(docTarget.getPosOfParagraph(targetPr));

                    if(npar <= pos) {
                        docTarget.setParagraph(srcPr, npar);
                    }

                }
                cursor = targetPr.getCTP().newCursor();

            } else if (elementType == BodyElementType.TABLE) {
                if(!bStart) {
                    continue;
                }
                if(bEnd) {
                    break;
                }

                XWPFTable table = (XWPFTable) bodyElement;
                String strid = table.getStyleID();
                if(strid != null) {
                    copyStyle(docSource, docTarget, docSource.getStyles().getStyle(strid));
                }

                XWPFTable dstTb = docTarget.insertNewTbl(cursor);
                int pos = docTarget.getTablePos(docTarget.getPosOfTable(dstTb));

                int npos = docTarget.getTables().size();
                if(npos >= pos) {
                    docTarget.setTable(pos, table);
                }
                cursor = dstTb.getCTTbl().newCursor();
            }

        }

        OutputStream outFile = new FileOutputStream("C:/Users/USER/Desktop/444.docx");

        docTarget.write(outFile);
        outFile.close();
        srcFile.close();
        targetFile.close();

    }

    /**复制文本格式
     * @param srcDoc
     * @param destDoc
     * @param style
     */
    private static void copyStyle(XWPFDocument srcDoc, XWPFDocument destDoc, XWPFStyle style) {
        if (destDoc == null || style == null) {
            return;
        }

        if (destDoc.getStyles() == null) {
            destDoc.createStyles();
        }

        List<XWPFStyle> usedStyleList = srcDoc.getStyles().getUsedStyleList(style);
        for (XWPFStyle xwpfStyle : usedStyleList) {
            destDoc.getStyles().addStyle(xwpfStyle);
        }
    }

    private static void createPic(String blipId, int id, long width, long height, CTInline inline) {
        String picXml = "" +
                "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                "   <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                "      <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                " <pic:nvPicPr>" +
                "    <pic:cNvPr id=\"" + id + "\" name=\"Generated\"/>" +
                "    <pic:cNvPicPr/>" +
                " </pic:nvPicPr>" +
                " <pic:blipFill>" +
                "    <a:blip r:embed=\"" + blipId + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>" +
                "    <a:stretch>" +
                "       <a:fillRect/>" +
                "    </a:stretch>" +
                " </pic:blipFill>" +
                " <pic:spPr>" +
                "    <a:xfrm>" +
                "       <a:off x=\"0\" y=\"0\"/>" +
                "       <a:ext cx=\"" + width + "\" cy=\"" + height + "\"/>" +
                "    </a:xfrm>" +
                "    <a:prstGeom prst=\"rect\">" +
                "       <a:avLst/>" +
                "    </a:prstGeom>" +
                " </pic:spPr>" +
                "      </pic:pic>" +
                "   </a:graphicData>" +
                "</a:graphic>";


        XmlToken xmlToken = null;
        try
        {
            xmlToken = XmlToken.Factory.parse(picXml);
        }
        catch(XmlException xe)
        {
            xe.printStackTrace();
        }
        inline.set(xmlToken);
        inline.setDistT(0);
        inline.setDistB(0);
        inline.setDistL(0);
        inline.setDistR(0);

        CTPositiveSize2D extent = inline.addNewExtent();
        extent.setCx(width);
        extent.setCy(height);

        CTNonVisualDrawingProps docPr = inline.addNewDocPr();
        docPr.setId(id);
        docPr.setName("Picture " + id);
        docPr.setDescr("Generated");
    }

}
