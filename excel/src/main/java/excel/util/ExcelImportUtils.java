package excel.util;


import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.File;
import java.io.FileInputStream;

import java.io.IOException;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;
import java.util.logging.Logger;

public class ExcelImportUtils {

    private final static String STRING_TYPE = "java.lang.String";
    private static Logger log = Logger.getAnonymousLogger();
    private final static String BOOLEAN_TYPE = "java.lang.Boolean";

    private static void check(String path, Object dataEntry) throws Exception {
        boolean exists = new File(path).exists();
        if (!exists)
            throw new IOException("不存在该路径上的文件");
        if (dataEntry == null)
            throw new RuntimeException("没有传递实体类");
    }

    /**
     *
     * @param filePath   文件所在位置
     * @param dataEntry  实体类（需要有一个无参构造,且要满足驼峰命名）
     * @param isReceiveEmpty  是否接受文件里面的空值
     * @param <T>
     * @return
     * @throws Exception  若isReceiveEmpty为 false，又出现空值，则会报数据无法接收,
     * 数据类型不对，会报类型对应不上，字段数量不一致会报错
     */
    public static <T> List<T> importExcel(String filePath, T dataEntry,boolean isReceiveEmpty) throws Exception {
        //校验
        check(filePath, dataEntry);

        //获取实体类字段
        Map<String, String> columns = getColumns(dataEntry);

        //获取excel中的字段即数据
        FileInputStream fileIn = new FileInputStream(filePath);

        //判断是xls还是xlsx
        Sheet sht0=null;
        if("xls".equals(filePath.split("\\.")[1])){
            HSSFWorkbook wb0 = new HSSFWorkbook(fileIn);//针对xls
            sht0 = wb0.getSheetAt(0);
        }else{
            XSSFWorkbook wb0=new XSSFWorkbook(fileIn);
            sht0 = wb0.getSheetAt(0);
        }



        String[] orderColumn = isColumnSame(columns, sht0.getRow(0));
        List<T> entryList = new ArrayList<T>();
        importData2Entry(entryList, sht0, orderColumn, dataEntry,isReceiveEmpty);
        fileIn.close();
        log.info("excel导入完成");
        return entryList;
    }

    private static <T> List<T> importData2Entry(List<T> entryList, Sheet sht0, String[] orderColumn, T dataEntry,boolean isReceiveNull) throws InvocationTargetException, IllegalAccessException, InstantiationException, ClassNotFoundException, NoSuchMethodException, NoSuchFieldException {
        int line = 0;

        for (Row r : sht0) {
            //去掉首行的字段名称
            if (line < 1) {
                line++;
                continue;
            }
            Class<?> aClass = dataEntry.getClass();
            String name = aClass.getName();
            T o = null;
            try {
                o = (T) aClass.getConstructor().newInstance(null);
            } catch (NoSuchMethodException e) {
                throw new NoSuchMethodException(name + " 需要有一个无参构造函数");
            }
            for (int i = 0; i < orderColumn.length; i++) {
                Field field = aClass.getDeclaredField(orderColumn[i]);
                field.setAccessible(true);//允许访问私有属性
                Class<?> type = field.getType();
                String methodName = "set" + orderColumn[i].substring(0, 1).toUpperCase() + orderColumn[i].substring(1, orderColumn[i].length());
                Method declaredMethod = aClass.getDeclaredMethod(methodName, type);
                try {
                    declaredMethod.invoke(o, setValueType(type, r.getCell(i),isReceiveNull));
                }catch (RuntimeException e){
                    throw new RuntimeException("第"+(i+1)+"列字段数据类型没对上");
                }

            }

            entryList.add(o);
            line++;
        }
        return null;
    }


    private static Object setValueType(Class<?> type, Cell cell,boolean isReceiveEmpty) throws RuntimeException {
        Object result = null;
        String packageName = type.getName();
        if (packageName.equals(STRING_TYPE)) {
            if(cell==null&&isReceiveEmpty){
                return null;
            }
            cell.setCellType(Cell.CELL_TYPE_STRING);
            result = cell.getStringCellValue();
            if((result==null||"".equals(result))&&isReceiveEmpty){
                return null;
            }
        } else if (packageName.equals(BOOLEAN_TYPE) || packageName.equals("boolean")) {
            if(cell==null&&isReceiveEmpty){
                return false;
            }
            cell.setCellType(Cell.CELL_TYPE_BOOLEAN);
            result = cell.getBooleanCellValue();
            if((result==null||"".equals(result))&&isReceiveEmpty){
                return false;
            }
        } else {
            //检验xls
            if(cell==null&&isReceiveEmpty){
                return 0;
            }
            cell.setCellType(Cell.CELL_TYPE_STRING);
            result=cell.getStringCellValue();
            //检验xlsx
            if((result==null||"".equals(result))&&isReceiveEmpty){
                return 0;
            }
            if((result+"").contains(".")){
                result=Float.parseFloat(result+"");
            }else{
                result=Integer.parseInt(result+"");
            }
        }

        return result;
    }

    //excel中的字段名和数量是否和实体类相同
    private static String[] isColumnSame(Map<String, String> columns, Row columnRow) {
        if (columns.size() != columnRow.getLastCellNum())
            throw new RuntimeException("字段数量对不上");
        String[] orderColumn = new String[columns.size()];

        for (int i = 0; i < columnRow.getLastCellNum(); i++) {
            columnRow.getCell(i).setCellType(Cell.CELL_TYPE_STRING);
            String stringCellValue = columnRow.getCell(i).getStringCellValue();
            orderColumn[i] = stringCellValue;
            columns.remove(stringCellValue);
        }
        if (columns.size() > 0)
            throw new RuntimeException("字段名称对不上");
        return orderColumn;
    }

    private static <T> Map<String, String> getColumns(T dataEntry) throws RuntimeException {
        Class dataEntryClass = dataEntry.getClass();
        int length = dataEntryClass.getDeclaredFields().length;
        if (length <= 0)
            throw new RuntimeException("该实体类中没有成员属性");
        Field[] declaredFields = dataEntryClass.getDeclaredFields();
        Map columns = new HashMap();
        for (int i = 0; i < declaredFields.length; i++) {
            String name = declaredFields[i].getName();
            columns.put(name, name);
        }
        return columns;
    }




}
