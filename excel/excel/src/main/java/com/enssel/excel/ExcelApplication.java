package com.enssel.excel;

import com.enssel.excel.file.XTX;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

@SpringBootApplication
public class ExcelApplication {

    public static void main(String[] args) {
//		SpringApplication.run(ExcelApplication.class, args);

		XTX xtx = new XTX("D:\\통합 문서 1(원본 삭제).xlsx");

        String[][] data = xtx.getData(0);

        for(String[] s1 : data){
            for(String s2 : s1){
                System.out.print(s2 + ", ");
            }
            System.out.println();
        }


    }

}
