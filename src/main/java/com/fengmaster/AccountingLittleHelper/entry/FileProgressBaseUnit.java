package com.fengmaster.AccountingLittleHelper.entry;

import com.fengmaster.AccountingLittleHelper.funcModule.AbsWordFuncModule;
import lombok.Data;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.util.List;

/**
 *
 * 文件处理基本单元
 * Created by Feng-master on 20/03/02.
 */
@Data
public class FileProgressBaseUnit {

    /**
     * 输出路径
     */
    private String outputFilePath;

    /**
     * 输入路径
     */
    private String inputFilePath;


    private XWPFDocument doc;



    /**
     * 行
     */
    private Row row;

}
