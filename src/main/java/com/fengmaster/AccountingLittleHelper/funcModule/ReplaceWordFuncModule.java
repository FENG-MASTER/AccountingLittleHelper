package com.fengmaster.AccountingLittleHelper.funcModule;

import com.fengmaster.AccountingLittleHelper.entry.FileProgressBaseUnit;
import com.fengmaster.AccountingLittleHelper.util.PoiUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.function.BiConsumer;

/**
 * Created by Feng-master on 20/03/02.
 */
public class ReplaceWordFuncModule extends AbsWordFuncModule {

    /**
     * key-列 val-替换的文本表达
     */
    private Map<Integer,String> indexRegMap=new HashMap<Integer, String>();


    public ReplaceWordFuncModule(Workbook workbook) {
        super(workbook);
        readReplaceSetting(workbook);
    }

    private void readReplaceSetting(Workbook workbook){
        Sheet sheet = workbook.getSheet("原始数据");
        Row settingRow = sheet.getRow(1);
        for (Cell cell : settingRow) {
//            {}这种表达将会被替换
            if (PoiUtil.getCellValue(cell).startsWith("｛")&&PoiUtil.getCellValue(cell).endsWith("｝")){
                //替换表达所在列
                int replaceRegColumnIndex=cell.getColumnIndex();
                //要被替换的表达式文本
                String replaceReg=PoiUtil.getCellValue(cell);
                indexRegMap.put(replaceRegColumnIndex,replaceReg);
            }

        }

    }

    @Override
    public void progress(FileProgressBaseUnit fileProgressBaseUnit) {
        Row row = fileProgressBaseUnit.getRow();

        Map<String, String> params=new HashMap<>();

        indexRegMap.forEach(new BiConsumer<Integer, String>() {
            @Override
            public void accept(Integer columnIndex, String reg) {
                Cell newTextCell = row.getCell(columnIndex);
                params.put(reg,PoiUtil.getCellValue(newTextCell));

            }
        });

        PoiUtil.wordReplace(fileProgressBaseUnit.getDoc(),params);

    }
}
