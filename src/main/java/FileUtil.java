import java.io.File;

public class FileUtil {
    /**
     * 校验文件是否合法
     *
     * @param fileUrl 文件路径（包含文件名）
     * @return 合法：true
     */
    static boolean isValidFile(String fileUrl) {
        if (fileUrl == null) return false;
        File file = new File(fileUrl);
        if (!file.isFile() || !file.exists()) {
            System.err.println("文件不存在，请检查文件路径是否正确");
            return false;
        }
        String[] format = file.getName().split("\\.");
        int suffixIndex = format.length - 1;
        if (!"xls".equals(format[suffixIndex]) && !"xlsx".equals(format[suffixIndex])) {
            System.err.println("输入文件名称为:" + format[0] + ". 请检查Excel文件格式，当前仅支持xls或xlsx格式");
            return false;
        }
        return true;
    }

    /**
     * 校验文件路径是否存在,如果不存在则尝试创建，不成功则返回false
     *
     * @param url 文件路径
     * @return 存在：true
     */
    static boolean validateOrCreate(String url) {
        if (url == null) return false;
        File file = new File(url);
        if (file.isFile()) {
            System.err.println("请输入文件路径，而非具体到文件 " + url);
            return false;
        }
        if (!file.isDirectory() && !file.mkdirs()) {
            System.err.println("文件路径不存在且无法创建，请检查路径" + url + "是否正确");
            return false;
        }
        return true;
    }

    /**
     * 返回文件名，如果不是文件返回空null
     *
     * @param url 文件路径
     * @return 文件名
     */
    static String getFileName(String url) {
        if (url == null) return null;
        File file = new File(url);
        return file.isFile() ? file.getName() : null;
    }
}
