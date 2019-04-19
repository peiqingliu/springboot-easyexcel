package com.springboot.easyexcel.utils;

import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.springboot.easyexcel.exception.ExcelException;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.List;

/**
 * easyexcel工具类
 * @auther liupeiqing
 * @date 2019/4/19 13:11
 */
public class ExcelUtil {

    /**
     * 读取 Excel(多个 sheet)
     * @param excel 文件
     * @param rowModel 实体类映射，继承 BaseRowModel 类
     * @return Excel 数据 list
     */
    public static List<Object> readExcel(MultipartFile excel, BaseRowModel rowModel){

        ExcelListener excelListener = new ExcelListener();
        ExcelReader reader = getReader(excel, excelListener);
        if (reader == null){
            return null;
        }

        for (Sheet sheet : reader.getSheets()){
            if (rowModel != null){
                sheet.setClazz(rowModel.getClass());
            }
            reader.read(sheet);
        }
        return excelListener.getDatas();
    }

    /**
     * 读取某个 sheet 的 Excel
     * @param excel excel
     * @param rowModel 实体类映射，继承 BaseRowModel 类
     * @param sheetNo sheet 的序号 从1开始
     * @return Excel 数据 list
     */
    public static List<Object> readExcel(MultipartFile excel,BaseRowModel rowModel,int sheetNo){
        return readExcel(excel,rowModel,sheetNo,1);
    }
    /**
     * 读取某个 sheet 的 Excel
     *
     * @param excel       文件
     * @param rowModel    实体类映射，继承 BaseRowModel 类
     * @param sheetNo     sheet 的序号 从1开始
     * @param headLineNum 表头行数，默认为1
     * @return Excel 数据 list
     */
    public static List<Object> readExcel(MultipartFile excel, BaseRowModel rowModel, int sheetNo,
                                         int headLineNum) {
        ExcelListener excelListener = new ExcelListener();
        ExcelReader reader = getReader(excel, excelListener);
        if (reader == null) {
            return null;
        }
        reader.read(new Sheet(sheetNo, headLineNum, rowModel.getClass()));
        return excelListener.getDatas();
    }

    /**
     *  导出 Excel ：一个 sheet，带表头
     * @param response HttpServletResponse
     * @param list 数据 list，每个元素为一个 BaseRowModel
     * @param fileName 导出的文件名
     * @param sheetName 导入文件的 sheet 名
     * @param rowModel  映射实体类，Excel 模型
     */
    public static void writeExcel(HttpServletResponse response,List<? extends BaseRowModel> list,
                                  String fileName,String sheetName,BaseRowModel rowModel){
        ExcelWriter excelWriter = new ExcelWriter(getOutputStream(fileName,response), ExcelTypeEnum.XLSX);
        Sheet sheet = new Sheet(1,0,rowModel.getClass());
        sheet.setSheetName(sheetName);
        excelWriter.write(list,sheet);
        excelWriter.finish();

    }

    /**
     *  导出 Excel ：多个 sheet，带表头
     * @param response HttpServletResponse
     * @param list 数据 list，每个元素为一个 BaseRowModel
     * @param fileName 导出的文件名
     * @param sheetName 导入文件的 sheet 名
     * @param obj 映射实体类，Excel 模型
     * @return
     */
    public static ExcelWriterFactroy writeExcelWithSheets(HttpServletResponse response,List<? extends BaseRowModel> list,
                                                        String fileName,String sheetName,BaseRowModel obj){
        ExcelWriterFactroy writer = new ExcelWriterFactroy(getOutputStream(fileName,response),ExcelTypeEnum.XLSX);
        Sheet sheet = new Sheet(1,0,obj.getClass());
        sheet.setSheetName(sheetName);
        writer.write(list,sheet);
        return writer;
    }

    /**
     * 导出文件时为Writer生成OutputStream
     * @param fileName 文件名称
     * @param response HttpServletResponse
     * @return
     */
    private static OutputStream getOutputStream(String fileName,HttpServletResponse response){
        //创建本地文件
        String filePath = fileName + ".xlsx";
        File dbFile = new File(filePath);
        try{
            if (!dbFile.exists() || dbFile.isDirectory()){
                dbFile.createNewFile();
            }
            fileName = new String(filePath.getBytes(),"ISO-8859-1");
            response.addHeader("Content-Disposition", "filename=" + fileName);
            return response.getOutputStream();
        } catch (IOException e) {
            throw new ExcelException("创建文件失败！");
        }
    }

    /**
     * 返回 ExcelReader
     * @param excel 需要解析的 Excel 文件
     * @param listener 自己创建的解析器
     * @return
     */
    private static ExcelReader getReader(MultipartFile excel,ExcelListener listener){
        String fleName = excel.getOriginalFilename();

        if (fleName == null || (!fleName.toLowerCase().endsWith(".xls") && !fleName.toLowerCase().endsWith(".xlsx"))){
            throw new ExcelException("文件格式错误！");
        }
        InputStream inputStream;
        try {
            inputStream = new BufferedInputStream(excel.getInputStream());
            return new ExcelReader(inputStream,null,listener,false);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }
}
