package org.parser.powerpoint;

import java.io.FileInputStream;

import org.apache.poi.xslf.usermodel.XMLSlideShow;

public class PPTParserExample {

    public static void main(String[] args) throws Exception {

        FileInputStream fis = new FileInputStream(args[0]);
        try {
        	XMLSlideShow ppt = new XMLSlideShow(fis);
        	PPTParser pptParser = new PPTParserImpl();
        	pptParser.processTextContainer(ppt);
        	//pptParser.processTemplateContainer(ppt);
            fis.close();
        } catch (Exception e) {
        	e.printStackTrace();
        }
    }
}
