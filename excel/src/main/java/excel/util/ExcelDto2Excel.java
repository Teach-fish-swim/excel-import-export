package excel.util;

import org.apache.poi.hssf.usermodel.*;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.*;
import java.util.logging.Logger;


public class ExcelDto2Excel {

    private static final String DEFAULT_PATH="c:/excel/";

    private static final String SHEET_NAME="sheet1";
    private static Logger log = Logger.getAnonymousLogger();

    /**
     * @param dataEntryList the entry of data about export   not null
     * @param filePath      export path         had default value
     * @param fileName      excel name          not null
     * @param sheetName      sheet name         had default value
     * @param <T>
     *
     * @throws Exception
     */
    public static <T> void excelDto2Excel(List<T> dataEntryList, String filePath, String fileName,String sheetName) throws Exception {
        // 获取总列数

        int CountColumnNum = getAllCount(dataEntryList);
        // 创建Excel文档
        HSSFWorkbook hwb = new HSSFWorkbook();

        // sheet 对应一个工作页
        if(sheetName==null||"".equals(sheetName))
            sheetName=SHEET_NAME;

        HSSFSheet sheet = hwb.createSheet(sheetName);
        HSSFRow firstrow = sheet.createRow(0); // 下标为0的行开始
        HSSFCell[] firstcell = new HSSFCell[CountColumnNum];
        String[] names = createColumn(CountColumnNum, dataEntryList.get(0).getClass());
        for (int j = 0; j < CountColumnNum; j++) {
            firstcell[j] = firstrow.createCell(j);
            firstcell[j].setCellValue(new HSSFRichTextString(names[j]));
        }
        for (int i = 0; i < dataEntryList.size(); i++) {
            // 创建一行
            HSSFRow row = sheet.createRow(i + 1);

            Object o = dataEntryList.get(i);


            parseData(o, CountColumnNum, row, names);
        }
        // 创建文件输出流，准备输出电子表格
        boolean exists = new File(filePath).exists();
        if(!exists){
            boolean mkdir = new File(DEFAULT_PATH).mkdirs();
            if(!mkdir)
                throw new IOException("excel存储路径创建失败,请检查是否已经存在 '"+DEFAULT_PATH+"' 文件夹");
            filePath=DEFAULT_PATH;
        }
        OutputStream out = new FileOutputStream(filePath + fileName + ".xls");
        hwb.write(out);
        out.close();
        log.info("数据导出成功");
        if(!exists)
            log.info("您的excel存放的位置："+DEFAULT_PATH);


    }

    private static void parseData(Object mapExportEntity, int CountColumnNum, HSSFRow row, String[] names) throws Exception {

        Class<?> aClass = mapExportEntity.getClass();
        for (int i = 0; i < CountColumnNum; i++) {
            HSSFCell xh = row.createCell(i);
            String methodName = "get" + names[i].substring(0, 1).toUpperCase() + names[i].substring(1, names[i].length());
            Method declaredMethod = aClass.getDeclaredMethod(methodName);
            Object invoke = declaredMethod.invoke(mapExportEntity, null);
            xh.setCellValue(invoke + "");
        }

    }

    //没问题
    private static String[] createColumn(int countColumnNum, Class clazz) {
        Field[] declaredFields = clazz.getDeclaredFields();
        String[] columns = new String[countColumnNum];
        //临时存储字段名
        List<String> list = new ArrayList();
        for (int i = 0; i < declaredFields.length; i++) {
            String name = declaredFields[i].getName();
            //排除不需要的字段
//            if (jsonObject.get(name) != null && "0".equals(jsonObject.get(name) + ""))
//                continue;
            list.add(name);
        }
        for (int i = 0; i < countColumnNum; i++)
            columns[i] = list.get(i);
        return columns;
    }

    //获取属性的个数
    private static <T> int getAllCount(List<T> clazz) throws RuntimeException {
        if (clazz == null || clazz.size() == 0)
            throw new RuntimeException("没有数据");
        Class aClass = clazz.get(0).getClass();
        int length = aClass.getDeclaredFields().length;
        int excludeCount = 0;
 //       Set<Map.Entry<String, Object>> entries = jsonObject.entrySet();
//        if (entries != null) {
//            Iterator<Map.Entry<String, Object>> iterator = entries.iterator();
//            if (iterator != null) {
//                while (iterator.hasNext()) {
//                    Map.Entry<String, Object> next = iterator.next();
//                    Object value = next.getValue();
//                    //值若为0，那么将会排除
//                    if ("0".equals(value + ""))
//                        excludeCount++;
//                }
//            }
//        }
//        System.out.println(length - excludeCount);
        return length - excludeCount;
    }


}