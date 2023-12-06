package com.enssel.excel.rest;


import com.grapecity.documents.excel.IPivotCache;
import com.grapecity.documents.excel.IPivotFields;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@RestController
@RequestMapping("/excel")
public class ExcelDataController {

    /**
     * 엑샐 파일과 시트를 생성 후 간단한 시험 데이터를 삽입하는 메소드
     */
    @GetMapping("/test1")
    public void test1() {
        Workbook workbook = new Workbook();

        IWorksheet worksheet1 = workbook.getWorksheets().add();
        IWorksheet worksheet2 = workbook.getWorksheets().addAfter(worksheet1);

        worksheet1.setName("시트 1");
        worksheet2.setName("시트 2");

        Object otherData = new Object[][]{
                {"T540p", 12, 9850},
                {"T570", 5, 7460},
                {"Y460", 6, 5400},
                {"Y460F", 8, 6240},};
        worksheet1.getRange("A5:C10").setValue(otherData);


        workbook.save("ExcelDataExam_001.xlsx");

    }

    /**
     * test1에서 만든 .xlsx 파일을 기반으로 데이터를 추출
     */
    @GetMapping("/test2")
    public void test2() {
        System.out.println("test2 실행");

        Workbook workbook = new Workbook();
//        workbook.open("ExcelDataExam_001.xlsx");
        workbook.open("통합 문서 1.xlsx");

        IWorksheet worksheet1 = workbook.getWorksheets().get(0);

        Object[][] datas = (Object[][]) worksheet1.getRange("B4:D6").getValue();

//        for(Object[] o1 : datas){
//            for(Object o : o1){
//                System.out.print(o + " ,");
//            }
//            System.out.println();
//        }

        List<Person> personList = new ArrayList<>();

        for (Object[] o1 : datas) {
            Object[] d = (Object[]) o1;

            personList.add(new Person((String) d[0], (Double) d[1], (Double) d[2]));
        }

        for (Person p : personList) {
            System.out.println(p);
        }
    }

    @GetMapping("/test3")
    public void test3() {
        System.out.println("test3 실행");

        Workbook workbook = new Workbook();

        workbook.open("통합 문서 1.xlsx");

        IWorksheet worksheet1 = workbook.getWorksheets().get(0);

        Object[][] datas = (Object[][]) worksheet1.getRange("A2:K61").getValue();

        for (Object[] o1 : datas) {
            for (Object o2 : o1) {
                System.out.print(o2 + ", ");
            }
            System.out.println();
        }

//        List<User> userList = new ArrayList<>();
//
//        for (Object[] o1 : datas) {
//            userList.add(new User(o1));
//        }
//
//        for (User u : userList) {
//            System.out.println(u);
//        }

    }

    /**
     * 피벗테이블 데이터 테스트
     */
    @GetMapping("/test4")
    public void test4() {
        System.out.println("test4 실행");

        Workbook workbook = new Workbook();

//        workbook.getPivotCaches().getCount();

        workbook.open("통합 문서 1(원본 삭제).xlsx");
//        workbook.open("통합 문서 1.xlsx");

        IWorksheet worksheet1 = workbook.getWorksheets().get(1);

        IPivotCache pivotCache = workbook.getPivotCaches().get(0);

        Object[][] pivotData = (Object[][]) workbook.getPivotCaches().get(0).getSourceData().getValue();

//        worksheet1.getPivotTables().get(0).get

//        Object[][] data = (Object[][]) worksheet1.getRange("A2:K61").getValue();

//        Object[][] pivotData = (Object[][]) pivotCache.getSourceData().getValue();
//        // 피벗 자제 데이터
//        Object[][] pivotData = (Object[][]) worksheet1.getPivotTables().get(0).getTableRange1().getValue();
//        // 데이터만
//        Object[][] pivotData = (Object[][]) worksheet1.getPivotTables().get(0).getDataBodyRange().getValue();

//        for (Object[] o1 : data) {
//            for (Object o2 : o1) System.out.print(o2 + ", ");
//            System.out.println();
//        }

        for (Object[] o1 : pivotData) {
            for (Object o2 : o1) System.out.print(o2 + ", ");
            System.out.println();
        }

//        IPivotFields pivotFields = worksheet1.getPivotTables().get(0).getRowFields();

//        System.out.println(pivotFields.get(0).getSourceName());

        System.out.println("피벗 테이블의 수" + worksheet1.getPivotTables().getCount());

        // 피벗 테이블 수만 큼 출력
        for (int i = 0, n = worksheet1.getPivotTables().getCount(); i < n; i++) {
            IPivotFields pivotFields = worksheet1.getPivotTables().get(i).getDataFields();


            System.out.println((i + 1) + "번째 피벗 테이블의 값 컬럼 정보:");

            for (int j = 0, n2 = pivotFields.getCount(); j < n2; j++) {
                System.out.println(pivotFields.get(j).getSourceName());
            }
        }


//        System.out.println(pivotTable);

//        Object[][] datas = (Object[][]) worksheet1.getRange("A2:K61").getValue();
//
//        List<User> userList = new ArrayList<>();
//
//        for (Object[] o1 : datas) {
//            userList.add(new User(o1));
//        }
//
//        for (User u : userList) {
//            System.out.println(u);
//        }

    }

    @GetMapping("/test5")
    public void test5() throws IOException {

        String filePath = "D:\\통합 문서 1(원본 삭제).xlsx";
        String zipPath = "D:\\통합 문서 1(원본 삭제).zip";

        File excelFile = new File(filePath);
        File zipFile = new File(zipPath);

        try {
            FileInputStream input = new FileInputStream(excelFile);
            FileOutputStream output = new FileOutputStream(zipFile);

            byte[] buf = new byte[1024];

            int readData;

            while((readData = input.read(buf)) > 0){
                output.write(buf, 0, readData);
            }

            input.close();
            output.close();
        } catch (Exception e) {
            System.out.println("오류발생");
            e.printStackTrace();
        }

        System.out.println(zipFile.length());

    }
}

class Person {
    String name;
    Double age;
    Double weight;

    Person() {
    }

    Person(String name, Double age, Double weight) {
        this.name = name;
        this.age = age;
        this.weight = weight;
    }

    @Override
    public String toString() {
        return "Person{" +
                "name='" + name + '\'' +
                ", age=" + age +
                ", weight=" + weight +
                '}';
    }
}

class User {
    Long idx;
    String name;
    String sex;
    Integer age;
    String dept;
    String phone;
    String addr;
    String rank;
    Boolean marry;
    Integer family;
    String hobby;

    User() {
    }

    public User(Long idx, String name, String sex, Integer age, String dept, String phone, String addr, String rank, Boolean marri, Integer family, String hobby) {
        this.idx = idx;
        this.name = name;
        this.sex = sex;
        this.age = age;
        this.dept = dept;
        this.phone = phone;
        this.addr = addr;
        this.rank = rank;
        this.marry = marri;
        this.family = family;
        this.hobby = hobby;
    }

    /**
     * Object의 형태로 데이터가 들어올 경우 User를 만들어 주는 생성자
     *
     * @param object User 정보가 담긴 object
     */
    public User(Object[] object) {
        this.idx = ((Double) object[0]).longValue();
        this.name = (String) object[1];
        this.sex = (String) object[2];
        this.age = ((Double) object[3]).intValue();
        this.dept = (String) object[4];
        this.phone = (String) object[5];
        this.addr = (String) object[6];
        this.rank = (String) object[7];
//        this.marri = (Boolean) object[8];
        this.marry = "O".equals(object[8]);
        this.family = ((Double) object[9]).intValue();
        this.hobby = (String) object[10];
    }

    @Override
    public String toString() {
        return "User{" +
                "idx=" + idx +
                ", name='" + name + '\'' +
                ", sex='" + sex + '\'' +
                ", age=" + age +
                ", dept='" + dept + '\'' +
                ", phone='" + phone + '\'' +
                ", addr='" + addr + '\'' +
                ", rank='" + rank + '\'' +
                ", marry=" + marry +
                ", family=" + family +
                ", hobby='" + hobby + '\'' +
                '}';
    }
}