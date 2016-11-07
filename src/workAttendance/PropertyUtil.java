package workAttendance;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

public class PropertyUtil {

    private static Map<String, Properties> map     = new HashMap<String, Properties>();

    private static Map<String, Long>       timeMap = new HashMap<String, Long>();

    public static String get(String name, String key) {
        if (!name.endsWith(".properties")) {
            name += ".properties";
        }

        // 查找到文件
        String path = "/";
        File file = null;
        if (path != null) {
            path = path.replace("\\", "/");
            path = path.endsWith("/") ? path : path + "/";
            file = new File(path + name);

            if (map.containsKey(name) && (!file.exists() || (file.exists() && timeMap.get(name) != null && timeMap.get(name) == file.lastModified()))) {
                return map.get(name).getProperty(key);
            }
        }

        InputStream is = null;
        Properties prop = null;
        try {
            is = PropertyUtil.class.getResourceAsStream(name); // junit
            // is = getClass().getClassLoader().getResourceAsStream(name); //web
            is = is != null ? is : PropertyUtil.class.getClassLoader().getResourceAsStream(name);
            if (is == null && file != null) {
                is = new FileInputStream(file);
            }
            prop = new Properties();
            prop.load(is);
            is.close();
            is = null;
            map.put(name, prop);
            if (file != null && file.exists()) {
                timeMap.put(name, file.lastModified());
            }
            return prop.getProperty(key);
        } catch (IOException e) {
            e.printStackTrace();
            if (is != null) {
                try {
                    is.close();
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
                is = null;
            }
        }
        return null;
    }
}
