package com.example.excel.handle;

import com.example.excel.annotation.Excel;
import com.example.excel.enums.AlignType;
import com.example.excel.enums.OperationType;
import com.example.excel.util.Encodes;
import com.example.excel.util.Reflections;
import com.google.common.collect.Lists;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import javax.servlet.http.HttpServletResponse;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.*;

/**
 * @Author: Jax
 * @Email: guoxingyou@xjiye.com
 * @Date: 2017/12/15/11:03
 * @Desc:
 **/
@Slf4j
public class ExportHandle {

    /**
     * 工作薄对象
     */
    private SXSSFWorkbook wb;

    /**
     * 工作表对象
     */
    private Sheet sheet;

    /**
     * 样式列表
     */
    private Map<String, CellStyle> styles;

    /**
     * 当前行号
     */
    private int rowNum;

    /**
     * 注解列表（Object[]{ ExcelField, Field/Method }）
     */
    List<Object[]> annotationList = Lists.newArrayList();

    /**
     * 构造函数
     * @param title 表格标题，传“空值”，表示无标题
     * @param cls 实体对象，通过annotation.ExportField获取标题
     */
    public ExportHandle(String title, Class<?> cls){
        this(title, cls, OperationType.ONLY_EXPORT);
    }


    /**
     * 构造函数
     * @param title 表格标题，传“空值”，表示无标题
     * @param cls 实体对象，通过annotation.Excel获取标题
     * @param type 导出类型（1:导出模板；2：导出数据）
     * @param groups 分组值
     */
    public ExportHandle(String title, Class<?> cls, OperationType type, int... groups){
        // Get annotation field
        Field[] fs = cls.getDeclaredFields();
        for (Field f : fs){
            //获取字段上加的@Excel注解
            Excel e = f.getAnnotation(Excel.class);
            if (e != null && (e.type().equals(OperationType.BOTH) || e.type().equals(type))){
                //根据字段注解中配置的groups进行筛选
                if (groups != null && groups.length > 0){
                    if(groups.length > 1 ){
                        throw new RuntimeException("一次操作只能存在一个分组.");
                    }
                    for (int efg : e.groups()){
                        if (groups[0] == efg){
                            annotationList.add(new Object[]{e, f});
                            break;
                        }
                    }
                }else{
                    //若无group属性，则直接将字段和对应的注解加入到一个全局的注解链表中，用于之后进行统一的排序
                    annotationList.add(new Object[]{e, f});
                }
            }
        }
        // Get annotation method
        Method[] ms = cls.getDeclaredMethods();
        for (Method m : ms){
            //获取方法上的注解
            Excel e = m.getAnnotation(Excel.class);
            if (e != null && (e.type().equals(OperationType.BOTH) || e.type().equals(type))){
                if (groups!=null && groups.length>0){
                    if(groups.length > 1 ){
                        throw new RuntimeException("一次操作只能存在一个分组.");
                    }
                    for (int efg : e.groups()){
                        if (groups[0] == efg){
                            annotationList.add(new Object[]{e, m});
                            break;
                        }
                    }
                }else{
                    annotationList.add(new Object[]{e, m});
                }
            }
        }
        // 对字段进行排序 Field sorting
        Collections.sort(annotationList, new Comparator<Object[]>() {
            public int compare(Object[] o1, Object[] o2) {
                return new Integer(((Excel)o1[0]).index()).compareTo(
                        new Integer(((Excel)o2[0]).index()));
            };
        });
        // Initialize
        List<String> headerList = Lists.newArrayList();
        for (Object[] os : annotationList){
            //获取注解title属性值
            String t = ((Excel)os[0]).title();
            // 如果是导出，则去掉注释
            if (type.equals(OperationType.ONLY_EXPORT)){
                String[] ss = StringUtils.split(t, "**", 2);
                if (ss.length==2){
                    t = ss[0];
                }
            }
            //将字段名称保存在一个list中，交给初始化方法使用
            headerList.add(t);
        }
        initialize(title, headerList);
    }

    /**
     * 初始化函数
     * @param title 表格标题，传“空值”，表示无标题
     * @param headerList 表头列表
     */
    private void initialize(String title, List<String> headerList) {
        this.wb = new SXSSFWorkbook(500);
        this.sheet = wb.createSheet("Export");
        this.styles = createStyles(wb);
        // Create title
        if (StringUtils.isNotBlank(title)){
            Row titleRow = sheet.createRow(rowNum++);
            titleRow.setHeightInPoints(30);
            Cell titleCell = titleRow.createCell(0);
            titleCell.setCellStyle(styles.get("title"));
            titleCell.setCellValue(title);
            sheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(),
                    titleRow.getRowNum(), titleRow.getRowNum(), headerList.size()-1));
        }
        // Create header
        if (headerList == null){
            throw new RuntimeException("headerList not null!");
        }
        Row headerRow = sheet.createRow(rowNum++);
        headerRow.setHeightInPoints(16);
        for (int i = 0; i < headerList.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellStyle(styles.get("header"));
            String[] ss = StringUtils.split(headerList.get(i), "**", 2);
            if (ss.length==2){
                cell.setCellValue(ss[0]);
                Comment comment = this.sheet.createDrawingPatriarch().createCellComment(
                        new XSSFClientAnchor(0, 0, 0, 0, (short) 3, 3, (short) 5, 6));
                comment.setString(new XSSFRichTextString(ss[1]));
                cell.setCellComment(comment);
            }else{
                cell.setCellValue(headerList.get(i));
            }
            sheet.autoSizeColumn(i);
        }
        for (int i = 0; i < headerList.size(); i++) {
            int colWidth = sheet.getColumnWidth(i)*2;
            sheet.setColumnWidth(i, colWidth < 3000 ? 3000 : colWidth);
        }
        log.debug("Initialize success.");
    }

    /**
     * 创建表格样式
     * @param wb 工作薄对象
     * @return 样式列表
     */
    private Map<String, CellStyle> createStyles(Workbook wb) {
        Map<String, CellStyle> styles = new HashMap<String, CellStyle>();

        CellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        Font titleFont = wb.createFont();
        titleFont.setFontName("Arial");
        titleFont.setFontHeightInPoints((short) 16);
        titleFont.setBold(true);
//        titleFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        style.setFont(titleFont);
        styles.put("title", style);

        style = wb.createCellStyle();
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        Font dataFont = wb.createFont();
        dataFont.setFontName("Arial");
        dataFont.setFontHeightInPoints((short) 10);
        style.setFont(dataFont);
        styles.put("data", style);

        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
        style.setAlignment(HorizontalAlignment.LEFT);
        styles.put("data1", style);

        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
        style.setAlignment(HorizontalAlignment.CENTER);
        styles.put("data2", style);

        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
        style.setAlignment(HorizontalAlignment.RIGHT);
        styles.put("data3", style);

        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
//		style.setWrapText(true);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font headerFont = wb.createFont();
        headerFont.setFontName("Arial");
        headerFont.setFontHeightInPoints((short) 10);
//        headerFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        headerFont.setBold(true);
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(headerFont);
        styles.put("header", style);

        return styles;
    }

    /**
     * 添加一行
     * @return 行对象
     */
    public Row addRow(){
        return sheet.createRow(rowNum++);
    }

    /**
     * 添加一个单元格
     * @param row 添加的行
     * @param column 添加列号
     * @param val 添加值
     * @return 单元格对象
     */
    public Cell addCell(Row row, int column, Object val){
        return this.addCell(row, column, val, AlignType.AUTO, Class.class);
    }

    /**
     * 添加一个单元格
     * @param row 添加的行
     * @param column 添加列号
     * @param val 添加值
     * @param align 对齐方式（1：靠左；2：居中；3：靠右）
     * @return 单元格对象
     */
    public Cell addCell(Row row, int column, Object val, AlignType align, Class<?> fieldType){
        Cell cell = row.createCell(column);
        CellStyle style = styles.get("data"+(align.num() >= 1 && align.num() <= 3 ? align : ""));
        try {
            if (val == null){
                cell.setCellValue("");
            } else if (val instanceof String) {
                cell.setCellValue((String) val);
            } else if (val instanceof Integer) {
                cell.setCellValue((Integer) val);
            } else if (val instanceof Long) {
                cell.setCellValue((Long) val);
            } else if (val instanceof Double) {
                cell.setCellValue((Double) val);
            } else if (val instanceof Float) {
                cell.setCellValue((Float) val);
            } else if (val instanceof Date) {
                DataFormat format = wb.createDataFormat();
                style.setDataFormat(format.getFormat("yyyy-MM-dd"));
                cell.setCellValue((Date) val);
            } else {
                if (fieldType != Class.class){
                    cell.setCellValue((String)fieldType.getMethod("setValue", Object.class).invoke(null, val));
                }else{
                    cell.setCellValue((String)Class.forName(this.getClass().getName().replaceAll(this.getClass().getSimpleName(),
                            "fieldtype."+val.getClass().getSimpleName()+"Type")).getMethod("setValue", Object.class).invoke(null, val));
                }
            }
        } catch (Exception ex) {
            log.info("Set cell value ["+row.getRowNum()+","+column+"] error: " + ex.toString());
            cell.setCellValue(val.toString());
        }
        cell.setCellStyle(style);
        return cell;
    }

    /**
     * 添加数据（通过annotation.ExportField添加数据）
     * @return list 数据列表
     */
    public <E> ExportHandle setDataList(List<E> list){
        for (E e : list){
            int column = 0;
            Row row = this.addRow();
            StringBuilder sb = new StringBuilder();
            for (Object[] os : annotationList){
                Excel excel = (Excel)os[0];
                Object val = null;
                // Get entity value
                try{
                    if (StringUtils.isNotBlank(excel.value())){
                        val = Reflections.invokeGetter(excel, excel.value());
                    }else{
                        if (os[1] instanceof Field){
                            val = Reflections.invokeGetter(e, ((Field)os[1]).getName());
                        }else if (os[1] instanceof Method){
                            val = Reflections.invokeMethod(e, ((Method)os[1]).getName(), new Class[] {}, new Object[] {});
                        }
                    }
                }catch(Exception ex) {
                    // Failure to ignore
                    log.info(ex.toString());
                    val = "";
                }
                this.addCell(row, column++, val, excel.align(), excel.fieldType());
                sb.append(val + ", ");
            }
            log.debug("Write success: ["+row.getRowNum()+"] "+sb.toString());
        }
        return this;
    }

    /**
     * 输出数据流
     * @param os 输出数据流
     */
    public ExportHandle write(OutputStream os) throws IOException {
        wb.write(os);
        return this;
    }

    /**
     * 输出到客户端
     * @param fileName 输出文件名
     */
    public ExportHandle write(HttpServletResponse response, String fileName) throws IOException{
        response.reset();
        response.setContentType("application/octet-stream; charset=utf-8");
        response.setHeader("Content-Disposition", "attachment; filename="+ Encodes.urlEncode(fileName));
        write(response.getOutputStream());
        return this;
    }

    /**
     * 输出到文件
     * @param name 输出文件名
     */
    public ExportHandle writeFile(String name) throws FileNotFoundException, IOException{
        FileOutputStream os = new FileOutputStream(name);
        this.write(os);
        return this;
    }

    /**
     * 清理临时文件
     */
    public ExportHandle dispose(){
        wb.dispose();
        return this;
    }
}
