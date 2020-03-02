package com.fengmaster.AccountingLittleHelper.funcModule;

import com.fengmaster.AccountingLittleHelper.entry.FileProgressBaseUnit;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 * 功能模块接口
 * Created by Feng-master on 20/03/02.
 */
public abstract class AbsWordFuncModule {

    protected Workbook workbook;


    public AbsWordFuncModule(Workbook workbook){
        this.workbook=workbook;
    }

    public abstract void progress(FileProgressBaseUnit fileProgressBaseUnit);



}
