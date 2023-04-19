package com.alibaba.easyexcel.test.demo.read;

import java.util.List;
import java.util.Map;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.metadata.Cell;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.excel.util.ListUtils;
import com.alibaba.fastjson2.JSON;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

/**
 * 读取头
 *
 * @author Jiaju Zhuang
 */
@Slf4j
public class CellDataDemoHeadDataListener implements ReadListener<CellDataReadDemoData> {
    /**
     * 每隔5条存储数据库，实际使用中可以100条，然后清理list ，方便内存回收
     */
    private static final int BATCH_COUNT = 100;

    private List<CellDataReadDemoData> cachedDataList = ListUtils.newArrayListWithExpectedSize(BATCH_COUNT);

    @Override
    public void invoke(CellDataReadDemoData data, AnalysisContext context) {
        Map<Integer, Cell> cellMap = context.readSheetHolder().getCellMap();
        for(Cell cell:cellMap.values()) {
            ReadCellData rd = (ReadCellData) cell;
            XSSFCellStyle xssfCellStyle = rd.getDataFormatData().getXssfCellStyle();
            XSSFColor xssfColor = xssfCellStyle.getFillForegroundColorColor();
            XSSFColor fontColor = xssfCellStyle.getFont().getXSSFColor();
            if(fontColor != null) {
                System.out.println("font-row:" + rd.getRowIndex() + ";col:" + rd.getColumnIndex() +";val:" + rd.getStringValue() +";color:" + fontColor.getARGBHex());
            }
            if(xssfColor != null) {
                System.out.println("back-row:" + rd.getRowIndex() + ";col:" + rd.getColumnIndex() +";val:" + rd.getStringValue() +";color:" + xssfColor.getARGBHex());
            }
        }

        log.info("解析到一条数据:{}", JSON.toJSONString(data));
        if (cachedDataList.size() >= BATCH_COUNT) {
            saveData();
            cachedDataList = ListUtils.newArrayListWithExpectedSize(BATCH_COUNT);
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        saveData();
        log.info("所有数据解析完成！");
    }

    /**
     * 加上存储数据库
     */
    private void saveData() {
        log.info("{}条数据，开始存储数据库！", cachedDataList.size());
        log.info("存储数据库成功！");
    }
}
