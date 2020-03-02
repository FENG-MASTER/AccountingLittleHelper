package com.fengmaster.AccountingLittleHelper;

import com.fengmaster.AccountingLittleHelper.progress.IProgress;
import com.fengmaster.AccountingLittleHelper.progress.XZHGenerator;

import java.util.ArrayList;
import java.util.List;

public class Main {

    public static void main(String[] args){

        List<IProgress> progresses=new ArrayList<IProgress>();
        progresses.add(new XZHGenerator());

        for (IProgress progress : progresses) {
            if (progress.getName().equals(args[0])){
                progress.progress(args);
            }

        }



    }


}
