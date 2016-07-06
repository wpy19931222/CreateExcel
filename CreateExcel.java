package com.dev.sl;

/**
 * Created by wpy19931222 on 2016/7/3.
 */

import com.dev.beans.*;
import jxl.*;
import jxl.format.*;
import jxl.format.Border;
import jxl.write.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.*;

public class CreateExcel
{
    private File target_file;
    private OutputStream os;
    private jxl.write.WritableWorkbook wwb;

    public boolean Start(int type,Customer customer)
    {
        if(!OpenExcel())
        {
            return false;
        }
        if(type == 1)
        {
            if(!WriteExcel_type1(customer))
            {
                return false;
            }
        }
        else if(type == 2)
        {
            if(!WriteExcel_type2(customer))
            {
                return false;
            }
        }
        else
        {
            return false;
        }
        if(!CloseExcel())
        {
            return false;
        }
        return true;
    }

    private boolean OpenExcel()
    {
        target_file = new File("abc.xls");
        try
        {
            os = new FileOutputStream(target_file);
            wwb = Workbook.createWorkbook(os);
        }
        catch (Exception e)
        {
            return false;
        }
        return true;
    }

    private boolean CloseExcel()
    {
        try
        {
            wwb.write();
            wwb.close();
            os.close();
        }
        catch (Exception e)
        {
            return false;
        }
        return true;
    }

    private boolean WriteExcel_type1(Customer customer)
    {
        WritableFont wf = new WritableFont(WritableFont.ARIAL, 11,
                WritableFont.NO_BOLD, false, UnderlineStyle.NO_UNDERLINE,
                jxl.format.Colour.BLACK);

        try
        {
            WritableSheet ws = wwb.createSheet("第一页",0);

            ws.getSettings().setShowGridLines(false);
            ws.setColumnView(0,30);
            ws.setColumnView(1,30);
            ws.setColumnView(2,30);
            ws.setColumnView(3,30);

            WritableCellFormat wcf = new WritableCellFormat(wf);
            wcf.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN,jxl.format.Colour.BLACK);

            WritableCellFormat wcf_noborder = new WritableCellFormat(wf);

            WritableCellFormat wcf_centre = new WritableCellFormat(wf);
            wcf_centre.setAlignment(jxl.format.Alignment.CENTRE);
            wcf_centre.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN,jxl.format.Colour.BLACK);

            jxl.write.Label[] label = new jxl.write.Label[27];

            MaintainDevice maintainDevice = customer.getMaintainDevice();

            label[0] = new jxl.write.Label(0, 0, "***"+customer.getCusName()+"取机凭证***",wcf_centre);
            label[1] = new jxl.write.Label(0, 1, "接修日期",wcf);
            label[2] = new jxl.write.Label(1, 1, new SimpleDateFormat("yyyy-MM-dd").format(customer.getDeliverTime()),wcf);
            label[3] = new jxl.write.Label(2, 1, "维修编号",wcf);
            label[4] = new jxl.write.Label(3, 1, Long.toString(maintainDevice.getMaintainID()),wcf);
            label[5] = new jxl.write.Label(0, 2, "产品类型",wcf);
            label[6] = new jxl.write.Label(1, 2, maintainDevice.getProductTypetoString(),wcf);
            label[7] = new jxl.write.Label(2, 2, "机器品牌",wcf);
            label[8] = new jxl.write.Label(3, 2, maintainDevice.getMachineBrand(),wcf);
            label[9] = new jxl.write.Label(0, 3, "机器型号",wcf);
            label[10] = new jxl.write.Label(1, 3, maintainDevice.getMachineModel(),wcf);
            label[11] = new jxl.write.Label(2, 3, "系列号",wcf);
            label[12] = new jxl.write.Label(3, 3, maintainDevice.getSerialNumber(),wcf);
            label[13] = new jxl.write.Label(0, 4, "单位名称",wcf);
            label[14] = new jxl.write.Label(1, 4, customer.getUnitName(),wcf);
            label[15] = new jxl.write.Label(2, 4, "联系人",wcf);
            label[16] = new jxl.write.Label(3, 4, customer.getContacts(),wcf);
            label[17] = new jxl.write.Label(0, 5, "机器故障现象",wcf_centre);
            label[18] = new jxl.write.Label(0, 6, maintainDevice.getBugDetails(),wcf_centre);
            label[19] = new jxl.write.Label(0, 7, "缺少零件",wcf_centre);
            label[20] = new jxl.write.Label(2, 7, "随机附件",wcf_centre);
            label[21] = new jxl.write.Label(0, 8, maintainDevice.getMissedPart(),wcf_centre);
            label[22] = new jxl.write.Label(2, 8, maintainDevice.getPart(),wcf_centre);
            label[23] = new jxl.write.Label(0, 9, "接机签字：",wcf_noborder);
            label[24] = new jxl.write.Label(1, 9, "机主签字：",wcf_noborder);
            label[25] = new jxl.write.Label(2, 9, "打印时间：",wcf_noborder);
            label[26] = new jxl.write.Label(3, 9, new SimpleDateFormat("yyyy-MM-dd").format(new Date()),wcf_noborder);

            for(int i = 0;i < 27;i ++)
            {
                ws.addCell(label[i]);
            }

            ws.mergeCells(0,0,3,0);
            ws.mergeCells(0,5,3,5);
            ws.mergeCells(0,6,3,6);
            ws.mergeCells(0,7,1,7);
            ws.mergeCells(2,7,3,7);
            ws.mergeCells(0,8,1,8);
            ws.mergeCells(2,8,3,8);
        }
        catch (Exception e)
        {
            return false;
        }

        return true;
    }
    private boolean WriteExcel_type2(Customer customer)
    {
        WritableFont wf = new WritableFont(WritableFont.ARIAL, 11,
                WritableFont.NO_BOLD, false, UnderlineStyle.NO_UNDERLINE,
                jxl.format.Colour.BLACK);

        try
        {
            WritableSheet ws = wwb.createSheet("第一页", 0);

            ws.getSettings().setShowGridLines(false);
            ws.setColumnView(0, 30);
            ws.setColumnView(1, 30);
            ws.setColumnView(2, 30);
            ws.setColumnView(3, 30);

            WritableCellFormat wcf = new WritableCellFormat(wf);
            wcf.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN,jxl.format.Colour.BLACK);

            WritableCellFormat wcf_noborder = new WritableCellFormat(wf);

            WritableCellFormat wcf_centre = new WritableCellFormat(wf);
            wcf_centre.setAlignment(jxl.format.Alignment.CENTRE);
            wcf_centre.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN,jxl.format.Colour.BLACK);

            jxl.write.Label[] label = new jxl.write.Label[41];

            MaintainDevice maintainDevice = customer.getMaintainDevice();
            BalanceCheck balanceCheck = maintainDevice.getBalanceCheck();
            MaintainRecord maintainRecord = maintainDevice.getMaintainRecord();
            List<BackUpPartFlow> bc = maintainRecord.getBackUpPartFlows();

            label[0] = new jxl.write.Label(0, 0, "***"+customer.getCusName()+"结算清单***",wcf_centre);
            label[1] = new jxl.write.Label(0, 1, "接修单号",wcf);
            label[2] = new jxl.write.Label(1, 1, Long.toString(maintainDevice.getMaintainID()),wcf);
            label[3] = new jxl.write.Label(0, 2, "接修日期",wcf);
            label[4] = new jxl.write.Label(1, 2, new SimpleDateFormat("yyyy-MM-dd").format(customer.getDeliverTime()),wcf);
            label[5] = new jxl.write.Label(2, 2, "修复日期",wcf);
            label[6] = new jxl.write.Label(3, 2, new SimpleDateFormat("yyyy-MM-dd").format(new Date()),wcf);
            label[7] = new jxl.write.Label(0, 3, "产品类型",wcf);
            label[8] = new jxl.write.Label(1, 3, maintainDevice.getProductTypetoString(),wcf);
            label[9] = new jxl.write.Label(2, 3, "机器品牌",wcf);
            label[10] = new jxl.write.Label(3, 3, maintainDevice.getMachineBrand(),wcf);
            label[11] = new jxl.write.Label(0, 4, "机器型号",wcf);
            label[12] = new jxl.write.Label(1, 4, maintainDevice.getMachineModel(),wcf);
            label[13] = new jxl.write.Label(2, 4, "系列号",wcf);
            label[14] = new jxl.write.Label(3, 4, maintainDevice.getSerialNumber(),wcf);
            label[15] = new jxl.write.Label(0, 5, "单位名称",wcf);
            label[16] = new jxl.write.Label(1, 5, customer.getUnitName(),wcf);
            label[17] = new jxl.write.Label(2, 5, "联系人",wcf);
            label[18] = new jxl.write.Label(3, 5, customer.getContacts(),wcf);
            label[19] = new jxl.write.Label(0, 6, "合计金额",wcf);
            label[20] = new jxl.write.Label(1, 6, Long.toString(balanceCheck.getMaterialCost()+balanceCheck.getMaintainCost()),wcf);
            label[21] = new jxl.write.Label(2, 6, "修理费"+Long.toString(balanceCheck.getMaintainCost()),wcf);
            label[22] = new jxl.write.Label(3, 6, "材料费"+ Long.toString(balanceCheck.getMaterialCost()),wcf);
            label[23] = new jxl.write.Label(0, 7, "机器故障现象",wcf_centre);
            label[24] = new jxl.write.Label(0, 8, maintainDevice.getBugDetails(),wcf_centre);
            label[25] = new jxl.write.Label(0, 9, "保修承诺",wcf_centre);
            label[26] = new jxl.write.Label(2, 9, "注意事项",wcf_centre);
            label[27] = new jxl.write.Label(0, 10, balanceCheck.getWarrantyPromise(),wcf_centre);
            label[28] = new jxl.write.Label(2, 10, balanceCheck.getAttention(),wcf_centre);
            label[29] = new jxl.write.Label(0, 11, "部件名称",wcf);
            label[30] = new jxl.write.Label(1, 11, "型号",wcf);
            label[31] = new jxl.write.Label(2, 11, "数量",wcf);
            label[32] = new jxl.write.Label(3, 11, "单价",wcf);
            int j = 0;
            for(j = 0;j < bc.size();j ++)
            {
                label[32+4*j+1] = new jxl.write.Label(0,11+j+1,bc.get(j).getBkPartName(),wcf);
                label[32+4*j+2] = new jxl.write.Label(0,11+j+1,bc.get(j).getBkPartModel(),wcf);
                label[32+4*j+3] = new jxl.write.Label(0,11+j+1,Integer.toString(bc.get(j).getAmount()),wcf);
                label[32+4*j+4] = new jxl.write.Label(0,11+j+1,Integer.toString(bc.get(j).getPrice()),wcf);
            }
            label[37] = new jxl.write.Label(0, 12+j, "发货签字：",wcf_noborder);
            label[38] = new jxl.write.Label(1, 12+j, "机主签字：",wcf_noborder);
            label[39] = new jxl.write.Label(2, 12+j, "打印时间：",wcf_noborder);
            label[40] = new jxl.write.Label(3, 12+j, new SimpleDateFormat("yyyy-MM-dd").format(new Date()),wcf_noborder);

            for(int i = 0;i < 41;i ++)
            {
                ws.addCell(label[i]);
            }

            ws.mergeCells(0,0,3,0);
            ws.mergeCells(0,7,3,7);
            ws.mergeCells(0,8,3,8);
            ws.mergeCells(0,9,1,9);
            ws.mergeCells(2,9,3,9);
            ws.mergeCells(0,10,1,10);
            ws.mergeCells(2,10,3,10);
        }
        catch (Exception e)
        {
            return false;
        }
        return true;
    }
}
