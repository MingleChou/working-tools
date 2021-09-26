package com.it.jiemin.common;

import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.InputStreamResource;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

/**
 * @author JieminZhou
 * @version 1.0
 * @date 2020/10/19 16:42
 */
public class ExcOperator {
    /**
     * 同一excel的多个版本融合
     *
     * @return 融合后文件的字节数组
     */
    public static byte[] merge(InputStreamResource[] fileInputAsResource, int index, int endIdx) throws IOException {
        // 拿出多个Excel表格
        XSSFWorkbook[] workbooks = new XSSFWorkbook[fileInputAsResource.length];
        for(int i = 0; i < fileInputAsResource.length; i++){
            workbooks[i] = new XSSFWorkbook(fileInputAsResource[i].getInputStream());
        }

        for(int n = index; n <= endIdx; n++) {
            //第index个sheet
            XSSFSheet sheet = workbooks[0].getSheetAt(n);

            //根据表头（第一行）确定每一行需要遍历的范围
            int startCellNum = sheet.getRow(0).getFirstCellNum();
            int endCellNum = sheet.getRow(0).getLastCellNum();

            //定义cell复制策略，跨Excel复制单元格时，需要禁止复制单元格样式，不然报错
            CellCopyPolicy.Builder builder = new CellCopyPolicy.Builder();
            CellCopyPolicy cellCopyPolicy = builder.cellFormula(true).cellStyle(false).cellValue(true)
                    .condenseRows(false).copyHyperlink(true).mergedRegions(true)
                    .mergeHyperlink(false).rowHeight(false).build();

            //遍历所有Excel，逐单元格对比
            for (int i = 1; i < workbooks.length; i++) {
                XSSFSheet tmpSheet = workbooks[i].getSheetAt(n);
                for (int j = 0; j < sheet.getLastRowNum(); j++) {
                    XSSFRow row = sheet.getRow(j);
                    XSSFRow tmpRow = tmpSheet.getRow(j);
                    for (int k = startCellNum; k < endCellNum; k++) {
                        //已第一个Excel为模板，如果单元格无内容，而另一个Excel对应单元格有内容，则将单元格内容复制过来；
                        //如果第一个单元格已经有内容，则以这一单元格内容为准，不在查看另一个Excel对应单元格内容
//                    System.out.println(n + " " + j + "   " + k);
//                    System.out.println(sheet.getLastRowNum());
                        if ((row.getCell(k) == null || row.getCell(k).getRawValue() == null || row.getCell(k).getRawValue().equals(""))
                                && (tmpRow.getCell(k) != null && tmpRow.getCell(k).getRawValue() != null && !tmpRow.getCell(k).getRawValue().equals(""))) {
                            if (row.getCell(k) == null) {
                                row.createCell(k).copyCellFrom(tmpRow.getCell(k), cellCopyPolicy);
                            } else {
                                row.getCell(k).copyCellFrom(tmpRow.getCell(k), cellCopyPolicy);
                            }
                        }
                    }
                }
            }
        }

        //将文件内容转为二进制数组返回
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        workbooks[0].write(baos);
        return baos.toByteArray();
    }
}
