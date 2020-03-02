package com.fengmaster.AccountingLittleHelper.progress;

import com.fengmaster.AccountingLittleHelper.util.PoiUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

/**
 * 询证函生成器
 */
public class XZHGenerator implements  IProgress{



    //模版map
    private Map<String,String> templateMap =new HashMap<String, String>();



    public String getName() {
        return "XZH";
    }

    public boolean progress(String[] args) {

        Workbook workbook=null;

        templateMap =new HashMap<String, String>();



        try {
            workbook = readWordBook(args[1]);
        } catch (IOException e) {
            e.printStackTrace();
        }

        templateMap= readConf(workbook);



        return false;
    }

    private List<Map<String,String>> readReplaceSetting(Workbook workbook){
        List<List<Map<String,String>> > allReplaceList=new LinkedList<List<Map<String, String>>>();
        Sheet sheet = workbook.getSheet("原始数据");

        Row settingRow = sheet.getRow(2);
        for (Cell cell : settingRow) {
//            {}这种表达将会被替换
            if (PoiUtil.getCellValue(cell).startsWith("{")&&PoiUtil.getCellValue(cell).endsWith("}")){

                //替换表达所在列
                int replaceRegColumnIndex=cell.getColumnIndex();
                //要被替换的表达式文本
                String replaceReg=PoiUtil.getCellValue(cell);

                int startRowIndex=3;

                Row replaceRow=sheet.getRow(startRowIndex);

                while (PoiUtil.isEmpty(replaceRow.getCell(1))){

                    Map<String,String> map=new HashMap<String, String>();
                    //需要替换的真实数据
                    Cell newText = replaceRow.getCell(replaceRegColumnIndex);
                    map.put(replaceReg,PoiUtil.getCellValue(newText));



                    replaceRow=sheet.getRow(++startRowIndex);
                }


            }

        }

        return allReplaceList;
    }

    private Workbook readWordBook(String filePath) throws IOException {
        Workbook workbook=null;
        InputStream fis = null;

        

        fis = new FileInputStream(filePath);
        if (filePath.endsWith(".xlsx")) {
            workbook = new XSSFWorkbook(fis);
        } else if (filePath.endsWith(".xls") || filePath.endsWith(".et")) {
            workbook = new HSSFWorkbook(fis);
        }
        fis.close();

        return workbook;

        }


    private Map<String,String> readConf(Workbook workbook){

        Map<String,String> tMap=new HashMap<String, String>();

        InputStream fis = null;
        try {
            /* 读EXCEL文字内容 */
            // 获取第一个sheet表，也可使用sheet表名获取
            Sheet sheet = workbook.getSheet("模版配置");
            // 获取行
            Iterator<Row> rows = sheet.rowIterator();
            Row row;
            Cell cell;
            rows.next();
            while (rows.hasNext()) {
                row = rows.next();
                // 获取单元格
                tMap.put(PoiUtil.getCellValue(row.getCell(0)),PoiUtil.getCellValue(row.getCell(1)));

            }
        } finally {
            if (null != fis) {
                try {
                    fis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }


        return tMap;

    }


}
