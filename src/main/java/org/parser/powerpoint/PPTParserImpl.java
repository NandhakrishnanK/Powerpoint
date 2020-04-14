package org.parser.powerpoint;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFBackground;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class PPTParserImpl implements PPTParser {

    public List<XSLFTextRun> processTextContainer(XMLSlideShow ppt) {
    	List<XSLFTextRun> textList = new ArrayList<XSLFTextRun>();
        for (XSLFSlide slide : ppt.getSlides()) {
            System.out.println("Title: " + slide.getTitle());

            for (XSLFShape shape : slide.getShapes()) {
                if (shape instanceof XSLFTextShape) {
                    XSLFTextShape tsh = (XSLFTextShape) shape;
                    for (XSLFTextParagraph p : tsh) {
                        for (XSLFTextRun r : p) {
                            System.out.println(r.getText());
                            textList.add(r);
                            try {
                            	System.out.println("  bold: " + r.isBold());
                            } catch(IllegalArgumentException ie) {
                            	System.out.println("  bold: false");
                            }
                            try {
                                System.out.println("  italic: " + r.isItalic());
                            } catch(IllegalArgumentException ie) {
                                System.out.println("  italic: false");
                            }
                            try {
                                System.out.println("  underline: " + r.isUnderline());
                            } catch(IllegalArgumentException ie) {
                                System.out.println("  underline: false");
                            }
                            try {
                            	System.out.println("  font.family: " + r.getFontFamily());
                            } catch(IllegalArgumentException ie) {
                            	System.out.println("  font.family: false");
                            }
                            try {
                                System.out.println("  font.size: " + r.getFontSize());
                            } catch(IllegalArgumentException ie) {
                                System.out.println("  font.size: false");
                            }
                            try {
                            	System.out.println("  font.color: " + r.getFontColor());
                            } catch(IllegalArgumentException ie) {
                            	System.out.println("  font.color: no");
                            }
                        }
                    }
                }
            }
        }
        return textList;
    }

    public void processTemplateContainer(XMLSlideShow ppt) {
    	for(XSLFSlideMaster master : ppt.getSlideMasters()){
    	    for(XSLFSlideLayout layout : master.getSlideLayouts()){
    	        System.out.println(layout.getType());
    	    }
    	}
    	XSLFSlideMaster defaultMaster = ppt.getSlideMasters()[0];
    	XSLFSlideLayout contentLayout = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
    	XSLFBackground tableLayout = defaultMaster.getBackground();
    }

}
