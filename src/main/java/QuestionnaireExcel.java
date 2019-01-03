import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class QuestionnaireExcel {
    private static final int QUESTIONNAIRE_BORDER = 7;
    private final String fileUrl;
    private String saveUrl;

    //统计表格中一共有多少问卷题目可以作为标题，key对应表格titleKey，value对应表格titleName
    private Map<Integer, String> questionnaireMap;

    //统计同一个用户的所有问卷，key对应表格orderNum，value为相同orderNum对应的所有行
    private Map<Integer, List<Row>> userRowsMap;

    public QuestionnaireExcel(String fileUrl, String saveUrl) {
        this.fileUrl = fileUrl;
        this.saveUrl = saveUrl;
        questionnaireMap = new TreeMap<>();
        userRowsMap = new HashMap<>();
    }


    /**
     * 将修改后的Excel存储到本地
     *
     * @param workbook Workbook
     * @throws IOException io exception
     */
    void writeExcelToLocal(Workbook workbook) {
        System.out.println("正在保存数据……");
        processSaveFileName();
        //使用try-with-resource语法
        try (OutputStream out = new FileOutputStream(saveUrl)) {
            workbook.write(out);
        } catch (IOException e) {
            System.out.println("保存失败,请关闭正在打开的文件，然后重试");
            System.err.println(e.getMessage());
        }
    }

    /**
     * 处理保存文件名，如果输入的路径文件名则是asset_processed.xlsx，
     * 否则按输入文件名称保存
     */
    private void processSaveFileName() {
        String fileName = FileUtil.getNameIfFileValid(saveUrl);
        if (fileName == null) {
            saveUrl += "/asset_processed.xlsx";
        } else {
            String[] format = fileName.split("\\.");
            String suffix = format[format.length - 1];
            if (!"xls".equals(suffix) && !"xlsx".equals(suffix)) {
                saveUrl += ".xlsx";
            }
        }
    }

    /**
     * 设置处理后的Sheet信息
     *
     * @param sheet Sheet
     */
    void setProcessedSheet(Sheet sheet) {
        int position = 1;
        for (Map.Entry<Integer, List<Row>> entry : userRowsMap.entrySet()) {
            List<Row> rows = entry.getValue();
            Row row = createSheetRow(sheet, position);
            position = setBasicInfo(position, rows, row);
            setAnswers(rows, row);
        }
    }

    /**
     * 设置回答问卷用户的基本信息
     *
     * @param userNum 用户问卷序号，从1开始
     * @param rows    每个用户对应源Excel表中的所有问卷答案
     * @param row     以创建的Excel行
     * @return 返回下一个用户问卷对应的序号
     */
    private int setBasicInfo(int userNum, List<Row> rows, Row row) {
        Row originRow = rows.get(0);
        row.getCell(0).setCellValue(userNum++);
        if (originRow.getCell(2) != null) {
            row.getCell(1).setCellValue(originRow.getCell(2).getStringCellValue());
        }
        if (originRow.getCell(3) != null) {
            row.getCell(2).setCellValue(originRow.getCell(3).getStringCellValue());
        }
        if (originRow.getCell(4) != null) {
            row.getCell(3).setCellValue(originRow.getCell(4).getStringCellValue());
        }
        if (originRow.getCell(5) != null) {
            row.getCell(4).setCellValue(originRow.getCell(5).getStringCellValue());
        }
        if (originRow.getCell(6) != null) {
            row.getCell(5).setCellValue(originRow.getCell(6).getStringCellValue());
        }
        if (originRow.getCell(7) != null) {
            row.getCell(6).setCellValue(originRow.getCell(7).getStringCellValue());
        }
        return userNum;
    }

    /**
     * 设置用户回答问卷的答案
     *
     * @param rows 每个用户对应源Excel表中的所有问卷答案
     * @param row  以创建的Excel行
     */
    private void setAnswers(List<Row> rows, Row row) {
        for (Row r : rows) {
            String titleKey = r.getCell(10).getStringCellValue();
            String answer = r.getCell(12).getStringCellValue();
            if (titleKey == null || answer == null) {
                continue;
            }
            row.getCell(Integer.valueOf(titleKey) + QUESTIONNAIRE_BORDER - 1).setCellValue(answer);
        }
    }

    /**
     * 创建Excel行
     *
     * @param sheet    Sheet
     * @param position 创建Excel行的位置
     * @return Row
     */
    private Row createSheetRow(Sheet sheet, int position) {
        Row row = sheet.createRow(position);
        for (int i = 0; i < QUESTIONNAIRE_BORDER + questionnaireMap.size(); i++) {
            row.createCell(i);
        }
        return row;
    }

    /**
     * 支持xls和xlsx格式，分别使用HSSFWorkbook和XSSFWorkbook创建
     *
     * @return Workbook
     * @throws IOException            io exception
     * @throws InvalidFormatException format exception
     */
    Workbook getWorkbook() throws IOException, InvalidFormatException {
        File excel = new File(fileUrl);
        String[] format = fileUrl.split("\\.");
        int index = format.length - 1;
        if ("xls".equals(format[index])) {
            //使用try-with-resource语法
            try (InputStream is = new FileInputStream(fileUrl);
                 Workbook workbook = new HSSFWorkbook(is)) {
                return workbook;
            }
        } else {
            return new XSSFWorkbook(excel);
        }
    }

    /**
     * 统计表格中一共有多少问卷题目和同一个用户回答的所有问卷
     *
     * @param sheet 处理的表格
     */
    void statisticQuestionnairesAndUserRows(Sheet sheet) {
        int lastRowNum = sheet.getLastRowNum();
        for (int rowNum = 1; rowNum <= lastRowNum; rowNum++) {
            String process = getProcessPercentage(rowNum + 1, lastRowNum + 1);
            if (process != null) {
                System.out.println("已处理:" + process);
            }
            Row row = sheet.getRow(rowNum);
            String titleKey = row.getCell(10).getStringCellValue();
            String titleName = row.getCell(11).getStringCellValue();
            if (titleKey == null || titleName == null) {
                System.out.println("请注意第" + (rowNum + 1) + "行titleKey或titleName为空");
                continue;
            }
            questionnaireMap.put(Integer.valueOf(titleKey), titleName);
            int orderNum = Integer.valueOf(row.getCell(1).getStringCellValue());
            List<Row> rows = userRowsMap.getOrDefault(orderNum, new ArrayList<>());
            rows.add(row);
            userRowsMap.put(orderNum, rows);
        }
    }

    /**
     * 创建表格标题行
     *
     * @param sheet 处理的表格
     */
    void createSheetRowTitle(Sheet sheet) {
        Row row = sheet.createRow(0);
        row.createCell(0).setCellValue("序号");
        row.createCell(1).setCellValue("日期");
        row.createCell(2).setCellValue("时间");
        row.createCell(3).setCellValue("IP");
        row.createCell(4).setCellValue("地址");
        row.createCell(5).setCellValue("提交类型");
        row.createCell(6).setCellValue("访问方式");
        int count = 7;
        for (String title : questionnaireMap.values()) {
            row.createCell(count++).setCellValue(title);
        }
    }

    /**
     * 模拟处理进度
     *
     * @param now 当前处理行
     * @param all 总行数
     * @return 处理百分比
     */
    private static String getProcessPercentage(int now, int all) {
        if (now == (int) (all * 0.1)) {
            return "10%";
        }
        if (now == (int) (all * 0.2)) {
            return "20%";
        }
        if (now == (int) (all * 0.3)) {
            return "30%";
        }
        if (now == (int) (all * 0.4)) {
            return "40%";
        }
        if (now == (int) (all * 0.5)) {
            return "50%";
        }
        if (now == (int) (all * 0.6)) {
            return "60%";
        }
        if (now == (int) (all * 0.7)) {
            return "70%";
        }
        if (now == (int) (all * 0.8)) {
            return "80%";
        }
        if (now == (int) (all * 0.9)) {
            return "90%";
        }
        if (now == all) {
            return "100%";
        }
        return null;
    }
}
