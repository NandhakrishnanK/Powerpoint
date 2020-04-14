package org.parser.powerpoint;

import java.util.List;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFTextRun;

public interface PPTParser {

	List<XSLFTextRun> processTextContainer(XMLSlideShow ppt);

	void processTemplateContainer(XMLSlideShow ppt);
	
	//List<XSLFTextRun> processTableContainer(XMLSlideShow ppt);
	
	//List<XSLFTextRun> processChartContainer(XMLSlideShow ppt);
}
