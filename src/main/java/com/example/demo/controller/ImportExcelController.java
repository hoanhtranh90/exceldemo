package com.example.demo.controller;

import com.example.demo.model.Info;
import org.apache.commons.math3.stat.descriptive.summary.Product;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@RestController
@RequestMapping("data")
public class ImportExcelController {


    //import file truc tiep
//    @PostMapping
//    public ResponseEntity<List<Info>> importExcelFile(@RequestParam("file") MultipartFile files) throws IOException {
//        HttpStatus status = HttpStatus.OK;
//        List<Info> productList = new ArrayList<>();
//
//        XSSFWorkbook workbook = new XSSFWorkbook(files.getInputStream());
//        XSSFSheet worksheet = workbook.getSheetAt(0);
//
//        for (int index = 9; index < worksheet.getPhysicalNumberOfRows(); index++) {
//            if (index > 0) {
//                Info info = new Info();
//
//                XSSFRow row = worksheet.getRow(index);
//                info.setMaSv(String.valueOf(row.getCell(1)));
//                String s =  (row.getCell(2))+" "+(row.getCell(3));
//
//                info.setHoTen(s);
//                info.setClassName(String.valueOf(row.getCell(4)));
//                info.setDiemThi(String.valueOf(row.getCell(8)));
//                String hocPhan = String.valueOf(row.getCell(5))
//
//                        ;
//                info.setHocPhan(hocPhan);
//
//
//                productList.add(info);
//            }
//        }
//
//        return new ResponseEntity<>(productList, status);
//    }

    //import file tu may
    @GetMapping
    public ResponseEntity<List<Info>> importExcelFile() throws IOException {
        HttpStatus status = HttpStatus.OK;
        List<Info> productList = new ArrayList<>();


        XSSFWorkbook workbook = new XSSFWorkbook("D:\\abc\\bbb.xlsx");
//        XSSFSheet worksheet = workbook.getSheetAt(10);

        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {

            XSSFSheet worksheet = workbook.getSheetAt(i);

            for (int index = 9; index < worksheet.getPhysicalNumberOfRows(); index++) {
                if (index > 0) {
                    Info info = new Info();

                    XSSFRow row = worksheet.getRow(index);
                    info.setMaSv(String.valueOf(row.getCell(1)));
                    String s = (row.getCell(2)) + " " + (row.getCell(3));

                    info.setHoTen(s);
                    info.setClassName(String.valueOf(row.getCell(4)));
                    info.setDiemThi(String.valueOf((row.getCell(16))));
//                    info.setDiemTongKet((row.getCell(17).getRawValue()));

                    String hocPhan = String.valueOf(worksheet.getRow(9).getCell(5));
                    info.setHocPhan(hocPhan);


                    productList.add(info);
                }
            }


        }
        return new ResponseEntity<>(productList, status);
    }


        @GetMapping("/search")
        public ResponseEntity<List<Info>> findData (@RequestParam(value = "MaSv") String Masv) throws IOException {
            HttpStatus status = HttpStatus.OK;
            List<Info> productList = new ArrayList<>();
            XSSFWorkbook workbook = new XSSFWorkbook("D:\\abc\\bbb.xlsx");

//            XSSFWorkbook workbook = new XSSFWorkbook("C:\\Users\\admin\\OneDrive\\Desktop\\BANG-DIEM-HOC-PHAN-LEN-MANG-NGANH-DA-PHONG-TIEN-\\BANG DIEM HOC PHAN LEN MANG NGANH DA PHONG TIEN\\aaa.xlsx");

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {

                XSSFSheet worksheet = workbook.getSheetAt(i);


                for (int index = 9; index < worksheet.getPhysicalNumberOfRows(); index++) {
                    if (index > 0) {
                        Info info = new Info();

                        XSSFRow row = worksheet.getRow(index);

                        if (String.valueOf(row.getCell(1)).equals(Masv)) {
                            info.setMaSv(String.valueOf(row.getCell(1)));
                            String s = (row.getCell(2)) + " " + (row.getCell(3));
                            info.setHoTen(s);
                            info.setClassName(String.valueOf(row.getCell(4)));
                            info.setDiemThi(String.valueOf((row.getCell(16))));
//                            info.setDiemTongKet(String.valueOf((row.getCell(17).getNumericCellValue())));
                            String hocPhan = String.valueOf(row.getCell(5));
                            info.setHocPhan(hocPhan);


                            productList.add(info);
                        }
                    }
                }
            }
            return new ResponseEntity<>(productList, status);
        }







}