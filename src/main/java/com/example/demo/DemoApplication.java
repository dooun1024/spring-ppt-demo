package com.example.demo;

import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.*;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.*;
import java.util.List;

@SpringBootApplication
@RestController
public class DemoApplication {

	public static void main(String[] args) {
		SpringApplication.run(DemoApplication.class, args);
	}

	@GetMapping
	public List<String> hello() throws IOException {

		String fileName = new String("サンプル.xls".getBytes("MS932"), "ISO-8859-1");

		HSSFSheet sheet;
		HSSFCell cell;

		System.out.println("__XXXXXXXXXXXXXXXXasdfasdf");

		XMLSlideShow ppt = new XMLSlideShow();
		ppt.createSlide();

		XSLFSlideMaster defaultMaster = ppt.getSlideMasters().get(0);

		XSLFSlideLayout layout
				= defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
		XSLFSlide slide = ppt.createSlide(layout);

		XSLFTextShape titleShape = slide.getPlaceholder(0);
		XSLFTextShape contentShape = slide.getPlaceholder(1);

		// スライドをループ
		for (XSLFShape shape : slide.getShapes()) {
			if (shape instanceof XSLFAutoShape) {
				// this is a template placeholder
			}
		}

		// 画像編集
		byte[] pictureData = IOUtils.toByteArray(
				new FileInputStream("test_image.png"));

		XSLFPictureData pd
				= ppt.addPicture(pictureData, PictureData.PictureType.PNG);
		XSLFPictureShape picture = slide.createPicture(pd);


		// PPT出力
		FileOutputStream out = new FileOutputStream("powerpoint.pptx");
		ppt.write(out);
		out.close();

		return List.of("Hello","World","22");
	}
}
