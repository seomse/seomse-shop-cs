/*
 * Copyright (C) 2020 Seomse Inc.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */


package com.seomse.shop.cs;

import com.seomse.commons.utils.ExceptionUtil;
import com.seomse.commons.utils.FileUtil;
import com.seomse.poi.excel.ExcelGet;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.*;

/**
 * @author macle
 */
public class CsExcelStats {

    private static final Logger logger = LoggerFactory.getLogger(CsExcelStats.class);

    private final ExcelGet excelGet;

    private XSSFRow row;
    private XSSFWorkbook work = null;
    /**
     * 생성자
     */
    public CsExcelStats(){
        excelGet = new ExcelGet();
    }


    private final  Comparator<CsMonthStats> sortYm = (c1, c2) -> {
        try {
            return Long.compare(new SimpleDateFormat("yyyy-MM").parse(c1.ym).getTime(), new SimpleDateFormat("yyyy-MM").parse(c2.ym).getTime());
        } catch (ParseException e) {
            e.printStackTrace();
            return 0;
        }
    };

    public void stats(String excelPath){

        try{
            work = new XSSFWorkbook(new FileInputStream(excelPath));
            excelGet.setXSSFWorkbook(work);
            XSSFSheet sheet = work.getSheetAt(0);
            int rowCount = excelGet.getRowCount(sheet);

            //정렬순위는 반품,

            Map<String, CsMonthStats> monthMap = new HashMap<>();

            for(int rowIndex = 1 ; rowIndex < rowCount ; rowIndex++){

                row = sheet.getRow(rowIndex);
                XSSFCell dateCell = row.getCell(1);
                if(dateCell == null ){
                    break;
                }

                System.out.println(rowIndex);

                Date date = dateCell.getDateCellValue();

                if(date == null){
                    break;
                }
                String ym = new SimpleDateFormat("yyyy-MM").format(date);

                CsMonthStats csMonthStats = monthMap.get(ym);
                if(csMonthStats == null){
                    csMonthStats = new CsMonthStats();
                    csMonthStats.ym = ym;
                    monthMap.put(ym, csMonthStats);
                }
                int count;

                try {
                     count = Integer.parseInt(getCellValue(5));
                }catch(Exception e){
                    e.printStackTrace();
                    continue;
                }
                csMonthStats.count += count;

                String level1 = getCellValue(3);
                CsCount csCount = csMonthStats.csCountMap.get(level1);
                if(csCount == null){
                    csCount = new CsCount();
                    csCount.name = level1;
                    csMonthStats.csCountMap.put(level1, csCount);
                }
                csCount.count += count;

                String shop = getCellValue(2);
                NameCount nameCount = csCount.shopCountMap.get(shop);
                if(nameCount == null){
                    nameCount = new NameCount();
                    nameCount.name = shop;
                    csCount.shopCountMap.put(shop, nameCount);
                }

                nameCount.count += count;

                nameCount = csMonthStats.shopCountMap.get(shop);
                if(nameCount == null){
                    nameCount = new NameCount();
                    nameCount.name = shop;
                    csMonthStats.shopCountMap.put(shop, nameCount);
                }

                nameCount.count += count;
                String level2 = getCellValue(4);

                CsCount childCs = csCount.childCsMap.get(level2);
                if(childCs == null){
                    childCs = new CsCount();
                    childCs.name = level2;
                    csCount.childCsMap.put(level2, childCs);
                }
                childCs.count += count;

                nameCount = childCs.shopCountMap.get(shop);
                if(nameCount == null){
                    nameCount = new NameCount();
                    nameCount.name = shop;
                    childCs.shopCountMap.put(shop, nameCount);
                }
                nameCount.count += count;
            }
            CsMonthStats [] csMonthStatsArray = monthMap.values().toArray(new CsMonthStats[0]);
            Arrays.sort(csMonthStatsArray, sortYm);

            Map<String, CsColumn> csColumnMap = new HashMap<>();

            // CS 컬럼 인덱스 설정
            for(CsMonthStats csMonthStats : csMonthStatsArray){
                Map<String, CsCount> csCountMap = csMonthStats.csCountMap;

                Collection<CsCount> csColl =  csCountMap.values();

                for(CsCount csCount: csColl ){
                    CsColumn csColumn  = csColumnMap.get(csCount.name);
                    if(csColumn == null){
                        csColumn = new CsColumn();
                        csColumn.name = csCount.name;
                        csColumnMap.put(csCount.name, csColumn);
                    }
                    Collection<CsCount> childColl =  csCount.childCsMap.values();
                    for(CsCount child : childColl){
                        if(csColumn.childIndexMap.containsKey(child.name)){
                            continue;
                        }
                        csColumn.childIndexMap.put(child.name,0);
                    }
                }
            }


            List<CsColumn> columnList = new ArrayList<>();
            CsColumn seqCsColumn = csColumnMap.remove("반품");
            if(seqCsColumn != null){
                columnList.add(seqCsColumn);
            }
            seqCsColumn = csColumnMap.remove("교환");
            if(seqCsColumn != null){
                columnList.add(seqCsColumn);
            }

            seqCsColumn = csColumnMap.remove("기타");
            columnList.addAll(csColumnMap.values());

            if(seqCsColumn != null){
                columnList.add(seqCsColumn);
            }

            int colIndex = 4;

            //noinspection ForLoopReplaceableByForEach
            for (int i = 0; i <columnList.size() ; i++) {
                CsColumn csColumn = columnList.get(i);
                csColumn.startIndex = colIndex;
                csColumn.endIndex = colIndex + csColumn.childIndexMap.size()-1;

                String [] names = csColumn.childIndexMap.keySet().toArray(new String[0]);
                Arrays.sort(names);
                //noinspection ForLoopReplaceableByForEach
                for (int j = 0; j <names.length ; j++) {
                    csColumn.childIndexMap.put( names[j], colIndex);
                    colIndex++;
                }

                csColumnMap.put(csColumn.name, csColumn);
            }


            sheet = work.createSheet("월별통계");
            work.setSheetOrder(sheet.getSheetName(),1);
            XSSFRow row= sheet.createRow(1);
            XSSFCell cell = row.createCell(1);
            cell.setCellValue("년월");
            cell = row.createCell(2);
            cell.setCellValue("구분");
            cell = row.createCell(3);
            cell.setCellValue("건수");
            sheet.addMergedRegion(new CellRangeAddress(1,2,1,1));
            sheet.addMergedRegion(new CellRangeAddress(1,2,2,2));
            sheet.addMergedRegion(new CellRangeAddress(1,2,3,3));


            sheet.createRow(2);


            Collection<CsColumn> csColumnColl =  csColumnMap.values();
            for(CsColumn csColumn : csColumnColl){
                row = sheet.getRow(1);
                cell = row.createCell(csColumn.startIndex);
                cell.setCellValue(csColumn.name);


                if(csColumn.startIndex != csColumn.endIndex) {
                    sheet.addMergedRegion(new CellRangeAddress(1, 1, csColumn.startIndex, csColumn.endIndex));
                }
                row = sheet.getRow(2);

                Set<String> childSet = csColumn.childIndexMap.keySet();
                for(String childName : childSet){
                    cell = row.createCell(csColumn.childIndexMap.get(childName));
                    cell.setCellValue(childName);
                }
            }


            int rowIndex = 3;
            for(CsMonthStats csMonthStats : csMonthStatsArray){
                row= sheet.createRow(rowIndex);
                cell = row.createCell(1);
                cell.setCellValue(csMonthStats.ym);

                cell = row.createCell(2);
                cell.setCellValue("전체");
                sheet.addMergedRegion(new CellRangeAddress(rowIndex,rowIndex+1,2,2));
                cell = row.createCell(3);
                cell.setCellValue(csMonthStats.count);
                sheet.addMergedRegion(new CellRangeAddress(rowIndex,rowIndex+1,3,3));

                int shopIndex = rowIndex +2;

                Map<String, Integer> rowLineMap = new HashMap<>();

                NameCount [] nameCounts = csMonthStats.shopCountMap.values().toArray(new NameCount[0]);
                for(NameCount nameCount : nameCounts){
                    row = sheet.createRow(shopIndex);
                    cell = row.createCell(2);
                    sheet.addMergedRegion(new CellRangeAddress(shopIndex,shopIndex+1,2,2));

                    cell.setCellValue(nameCount.name);
                    cell = row.createCell(3);
                    sheet.addMergedRegion(new CellRangeAddress(shopIndex,shopIndex+1,3,3));
                    cell.setCellValue(nameCount.count);
                    rowLineMap.put(nameCount.name, shopIndex);

                    shopIndex++;
                    shopIndex++;
                }

                Collection<CsCount> csCounts = csMonthStats.csCountMap.values();
                sheet.createRow(rowIndex+1);
                for (CsCount csCount : csCounts) {

                    CsColumn csColumn = csColumnMap.get(csCount.name);

                    row = sheet.getRow(rowIndex);
                    cell = row.createCell(csColumn.startIndex);

                    if(csColumn.startIndex != csColumn.endIndex) {
                        sheet.addMergedRegion(new CellRangeAddress(rowIndex,rowIndex,csColumn.startIndex,csColumn.endIndex));
                    }
                    cell.setCellValue(csCount.count);



                    Collection<CsCount> childColl = csCount.childCsMap.values();
                    for(CsCount child : childColl){
                        int childColumn = csColumn.childIndexMap.get(child.name);
                        row = sheet.getRow(rowIndex+1);
                        row.createCell( childColumn).setCellValue(child.count);
                        Collection<NameCount> nameCountColl = child.shopCountMap.values();

                        for(NameCount nameCount : nameCountColl){
                            int line = rowLineMap.get(nameCount.name)+1;
                            row = sheet.getRow(line);
                            if(row == null){
                                row = sheet.createRow(line);
                            }
                            row.createCell( childColumn).setCellValue(nameCount.count);
                        }


                    }

                    Collection<NameCount> nameCountColl = csCount.shopCountMap.values();
                    for(NameCount nameCount : nameCountColl){
                        int line = rowLineMap.get(nameCount.name);
                        row = sheet.getRow(line);
                        cell = row.createCell(csColumn.startIndex);
                        cell.setCellValue(nameCount.count);
                        if(csColumn.startIndex != csColumn.endIndex) {
                            sheet.addMergedRegion(new CellRangeAddress(line,line,csColumn.startIndex,csColumn.endIndex));
                        }
                    }

                }

                sheet.addMergedRegion(new CellRangeAddress(rowIndex,shopIndex-1,1,1));
                sheet.addMergedRegion(new CellRangeAddress(shopIndex,shopIndex,1,colIndex-1));
                rowIndex = shopIndex+1;
            }


            sheet.setColumnWidth(0, 200);
            sheet.setColumnWidth(1, 2500);

            sheet.setColumnWidth(2, 3000);

            sheet.setColumnWidth(3, 1500);
            for(int i=4 ; i<colIndex ; i++){
                sheet.autoSizeColumn(i);
                sheet.setColumnWidth(i, sheet.getColumnWidth(i)+100);
            }

            //틀고정
            sheet.createFreezePane(1,3);

            Font headerFont = work.createFont();
            headerFont.setFontName("맑은 고딕");
            headerFont.setBold(true);

            CellStyle headStyle = work.createCellStyle();
            headStyle.setBorderTop(BorderStyle.THIN);
            headStyle.setBorderBottom(BorderStyle.THIN);
            headStyle.setBorderLeft(BorderStyle.THIN);
            headStyle.setBorderRight(BorderStyle.THIN);
            headStyle.setAlignment(HorizontalAlignment.CENTER);
            headStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            for(int i=1 ; i<rowIndex-1;i++){

                row = sheet.getRow(i);
                if(row == null){
                    row = sheet.createRow(i);
                }
                for(int j = 1 ; j<colIndex ; j++){
                    if(j == 1) {
                        cell = row.getCell(j);
                        if (cell == null) {
                            cell = row.createCell(j);
                        }

                        CellStyle dateStyle = makeStyle();
                        dateStyle.setFont(headerFont);
                        cell.setCellStyle(dateStyle);
                    }else{
                        cell = row.getCell(j);
                        if (cell == null) {
                            cell = row.createCell(j);
                        }
                        cell.setCellStyle(headStyle);
                    }
                }
            }

            //년월
            headStyle  = makeHeaderStyle();
            headStyle.setFont(headerFont);
            headStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.LIGHT_TURQUOISE.getIndex());
            XSSFCell headCell = sheet.getRow(1).getCell(1);
            headCell.setCellStyle(headStyle);

            //구분
            headStyle  = makeHeaderStyle();
            headStyle.setFont(headerFont);
            headStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.PALE_BLUE.getIndex());
            headCell = sheet.getRow(1).getCell(2);
            headCell.setCellStyle(headStyle);

            //문의수
            headStyle  = makeHeaderStyle();
            headStyle.setFont(headerFont);
            headStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.SKY_BLUE.getIndex());
            headCell = sheet.getRow(1).getCell(3);
            headCell.setCellStyle(headStyle);

            //반품
            headStyle  = makeHeaderStyle();
            headStyle.setFont(headerFont);
            headStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.ROSE.getIndex());
            headCell = sheet.getRow(1).getCell(4);
            headCell.setCellStyle(headStyle);

            //교환
            CsColumn csColumn= csColumnMap.get("교환");
            headStyle  = makeHeaderStyle();
            headStyle.setFont(headerFont);
            headStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.GOLD.getIndex());
            headCell = sheet.getRow(1).getCell(csColumn.startIndex);
            headCell.setCellStyle(headStyle);

            //배송
            csColumn= csColumnMap.get("배송");
            headStyle  = makeHeaderStyle();
            headerFont = work.createFont();
            headerFont.setBold(true);
            headerFont.setFontName("맑은 고딕");
            headerFont.setColor(IndexedColors.WHITE.getIndex());
            headStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.LIGHT_BLUE.getIndex());
            headStyle.setFont(headerFont);
            headCell = sheet.getRow(1).getCell(csColumn.startIndex);
            headCell.setCellStyle(headStyle);

            //기타
            csColumn= csColumnMap.get("기타");
            headStyle  = makeHeaderStyle();
            headStyle.setFont(headerFont);
            headStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.GREY_50_PERCENT.getIndex());
            headCell = sheet.getRow(1).getCell(csColumn.startIndex);
            headCell.setCellStyle(headStyle);

            sheet.setTabColor(new XSSFColor(Color.BLUE));
            String fileName = new File(excelPath).getParentFile().getAbsolutePath() + "/CS_통계_" + new SimpleDateFormat("yyyyMMdd").format(new Date()) +".xlsx";

            File file = new File(fileName);
            if(file.isFile()){
                fileName = FileUtil.makeName(file);
                file = new File(fileName);
            }

            FileOutputStream fos = null;
            try {

                fos = new FileOutputStream(file);
                work.write(fos);

            }catch(Exception e){
                e.printStackTrace();
            }finally {
                try {
                    if(fos!=null) fos.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }


        }catch(Exception e){
            logger.error(ExceptionUtil.getStackTrace(e));
        }finally {
            try {
                if(work!=null) work.close();
            } catch (IOException e) {
                logger.error(ExceptionUtil.getStackTrace(e));
            }
        }
    }

    private String getCellValue(int cellNum){
        return excelGet.getCellValue(row, cellNum);
    }

    public CellStyle makeStyle(){
        CellStyle cellStyle = work.createCellStyle();
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return cellStyle;

    }


    public CellStyle makeHeaderStyle(){
        CellStyle cellStyle = work.createCellStyle();
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return cellStyle;

    }


    public static void main(String[] args) {
        CsExcelStats csExcelStats = new CsExcelStats();
        csExcelStats.stats("C:\\Users\\macle\\Desktop\\CS_Stats.xlsx");
    }

}
