import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Main {
    public static void main(String[] args) {
        if (!validateParams(args)) {
            return;
        }
        System.out.println("数据正在处理,请稍等……");
        QuestionnaireExcel excel = new QuestionnaireExcel(args[0], args[1]);
        try {
            Workbook workbook = excel.getWorkbook();
            Sheet sheet1 = workbook.getSheetAt(0);
            Sheet sheet2 = workbook.createSheet("处理后");
            excel.statisticQuestionnairesAndUserRows(sheet1);
            excel.createSheetRowTitle(sheet2);
            excel.setProcessedSheet(sheet2);
            if (excel.saveToLocal(workbook, args)) {
                System.out.println("处理完成");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 校验文件路径是否有效
     *
     * @param args 输入文件路径
     * @return 有效：true
     */
    private static boolean validateParams(String[] args) {
        return isValidArgs(args) && FileUtil.isValidFile(args[0]) && FileUtil.validateOrCreate(args[1]);
    }

    /**
     * 校验输入参数个数
     *
     * @param args 输入文件路径
     * @return 有效：true
     */
    private static boolean isValidArgs(String[] args) {
        if (args == null || args.length < 2) {
            System.err.println("请输入文件所在路径和文件保存路径");
            return false;
        }
        return true;
    }
}
