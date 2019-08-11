package com.seomse.shop.cs;

import com.seomse.commons.file.FileUtil;
import com.seomse.commons.utils.ExcelGet;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.*;

/**
 * <pre>
 *  파 일 명 : CsExcelStats.java
 *  설    명 : Cs Excel  통계
 *
 *  작 성 자 : macle
 *  작 성 일 : 2019.08
 *  버    전 : 1.0
 *  수정이력 :
 *  기타사항 :
 * </pre>
 * @author Copyrights 2019 by ㈜섬세한사람들. All right reserved.
 */
public class CsExcelStats {

//    private static final Logger logger = LoggerFactory.getLogger(CsExcelStats.class);

    private ExcelGet excelGet;
    @SuppressWarnings("FieldCanBeLocal")
    private XSSFSheet sheet;
    private XSSFRow row;
    /**
     * 생성자
     */
    public CsExcelStats(){
        excelGet = new ExcelGet();
    }

    @SuppressWarnings("Convert2Lambda")
    private final  Comparator<CsMonthStats> sortYm =  new Comparator<CsMonthStats>() {
        @Override
        public int compare(CsMonthStats c1, CsMonthStats c2 ) {
            try {
                return Long.compare(new SimpleDateFormat("yyyy-MM").parse(c1.ym).getTime(),new SimpleDateFormat("yyyy-MM").parse(c2.ym).getTime());
            } catch (ParseException e) {
                e.printStackTrace();
                return 0;
            }
        }
    };

    public void stats(String excelPath){
        XSSFWorkbook work = null;
        try{
            work = new XSSFWorkbook(new FileInputStream(excelPath));
            excelGet.setXSSFWorkbook(work);
            sheet = work.getSheetAt(0);
            int rowCount = excelGet.getRowCount(sheet);

            //정렬순위는 반품,

            Map<String, CsMonthStats> monthMap = new HashMap<>();

            for(int rowIndex = 2 ; rowIndex < rowCount ; rowIndex++){

                row = sheet.getRow(rowIndex);
                XSSFCell dateCell = row.getCell(1);
                if(dateCell == null ){
                    break;
                }

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
            XSSFRow row= sheet.createRow(1);
            XSSFCell cell = row.createCell(1);
            cell.setCellValue("년월");
            cell = row.createCell(2);
            cell.setCellValue("구분");
            cell = row.createCell(3);
            cell.setCellValue("문의수");
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

            sheet.setColumnWidth(1, 2100);

            sheet.setColumnWidth(2, 3000);

            sheet.setColumnWidth(3, 1600);
            for(int i=4 ; i<colIndex ; i++){
                sheet.autoSizeColumn(i);
                sheet.setColumnWidth(i, sheet.getColumnWidth(i)+100);
            }


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
                    cell = row.getCell(j);
                    if(cell == null){
                        cell = row.createCell(j);
                    }
                    cell.setCellStyle(headStyle);
                }
            }

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
            e.printStackTrace();
        }finally {
            try {
                if(work!=null) work.close();
            } catch (IOException e) {
              e.printStackTrace();
            }
        }
    }

    private String getCellValue(int cellNum){
        return excelGet.getCellValue(row, cellNum);
    }


    public static void main(String[] args) {
        CsExcelStats csExcelStats = new CsExcelStats();
        csExcelStats.stats(args[0]);

//        csExcelStats.stats("C:\\Users\\macle\\Desktop\\CS_Stats.xlsx");
    }

}
