package org.example.compare.bean;

import java.util.List;

public class CompareFileConfig {
    private List<FileConfig> files;

    // Getter and setter for files

    public static class FileConfig {
        public String fileName;
        public String keyColumn;
        public int dataStartLineId;

        public String dataEndColumn;  // 文件最大列数,先支持
    }

    public List<FileConfig> getFiles() {
        return files;
    }

    public void setFiles(List<FileConfig> files) {
        this.files = files;
    }
}
