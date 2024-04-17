package org.example.compare.controller;

import cn.hutool.core.io.resource.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.core.io.UrlResource;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.net.MalformedURLException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

import org.example.compare.bean.CompareFileConfig;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

@Component

public class ExcelComparator {

    //  此处为mac下运行才会有的问题，因为mac会自动创建一个隐藏的macosx的副本文件,若为linux或者windows则没有这个问题。
     String specialKey = "MACOSX";

    public Map<String, List<CompareResult>> compareExcelFiles(
            String oriFolder,
            String tarFolder,
            Map<String, CompareFileConfig.FileConfig> configMap
    ) throws IOException {
        File[] oriFiles = new File(oriFolder).listFiles();
        File[] tarFiles = new File(tarFolder).listFiles();
            Map<String, List<CompareResult>> resMap = new HashMap<>();
            if (oriFiles != null) {
                for (File oriFile : oriFiles) {
                    List<CompareResult> excelList = new ArrayList<>();
                    int count = 0;
                    // 读不到配置项内容或者名字中有特殊key
                    if (oriFile.getName().contains(specialKey) || !configMap.containsKey(oriFile.getName())) {
                        continue;
                    }
                    if (isExcelFile(oriFile)) {
                        CompareResult resExcel = new CompareResult(oriFile.getName());
                        resExcel.isSame = false;
                        resExcel.errorType = CompareErrorType.NotExcelFile;
                        resExcel.isTarget = false;
                        excelList.add(resExcel);
                        continue;
                    }

                    if (tarFiles != null) {
                        for (File tarFile : tarFiles) {
                            if (tarFile.getName().contains(specialKey)||
                                    !configMap.containsKey(tarFile.getName())) {
                                continue;
                            }
                            if (isExcelFile(tarFile)) {
                                CompareResult resExcel = new CompareResult(tarFile.getName());
                                resExcel.isSame = false;
                                resExcel.errorType = CompareErrorType.NotExcelFile;
                                resExcel.isTarget = true;
                                excelList.add(resExcel);
                                continue;
                            }
                            if (oriFile.getName().equals(tarFile.getName())) {
                                List<Integer> keywordColList = new ArrayList<>();
                                int columnId = columnToNumber(configMap.get(oriFile.getName()).keyColumn);
                                keywordColList.add(columnId);
                                List<CompareResult> resExcelList = compareExcelSheets(oriFile, tarFile,
                                        keywordColList, configMap.get(oriFile.getName()).dataStartLineId);
                                excelList.addAll(resExcelList);
                                count++;
                                break;
                            }
                        }
                    }
                    if (count == 0) {
                        // 没有同名文件
                        CompareResult resExcel = new CompareResult(oriFile.getName());
                        resExcel.isSame = false;
                        resExcel.errorType = CompareErrorType.MissingFile;
                        excelList.add(resExcel);
                    }
                    if (excelList.size() > 0){
                        resMap.put(oriFile.getName(), excelList);
                    }
                }
                if (tarFiles != null) {
                    for (File tarFile : tarFiles) {
                        List<CompareResult> excelList = new ArrayList<>();
                        int count = 0;
                        if (tarFile.getName().contains(specialKey) || !configMap.containsKey(tarFile.getName())) {
                            continue;
                        }
                        if (isExcelFile(tarFile)) {
                            CompareResult resExcel = new CompareResult(tarFile.getName());
                            resExcel.isSame = false;
                            resExcel.errorType = CompareErrorType.NotExcelFile;
                            resExcel.isTarget = true;
                            excelList.add(resExcel);
                            continue;
                        }
                        for (File oriFile : oriFiles) {
                            if (oriFile.getName().equals(tarFile.getName())) {
                                count++;
                                break;
                            }
                        }
                        if (count == 0) {
                            CompareResult resExcel = new CompareResult(tarFile.getName());
                            resExcel.isSame = false;
                            resExcel.errorType = CompareErrorType.MissingFile;
                            resExcel.isTarget = true;
                            excelList.add(resExcel);
                        }
                        if (excelList.size() > 0){
                            resMap.put(tarFile.getName(), excelList);
                        }
                    }
                }
            }
            // 将resMap遍历，将每个v的list里的CompareResult按照isTarget排序
            for (Map.Entry<String, List<CompareResult>> entry : resMap.entrySet()) {
                entry.getValue().sort(Comparator.comparing(CompareResult::isTarget));
            }
        return resMap;
    }

    public  List<CompareResult> compareExcelSheets(
            File oriFile,
            File tarFile,
            List<Integer> keywordColList,
            int dataStartLine) throws IOException {
        try (XSSFWorkbook oriWorkbook = new XSSFWorkbook(new FileInputStream(oriFile));
             XSSFWorkbook tarWorkbook = new XSSFWorkbook(new FileInputStream(tarFile))) {
            List<CompareResult> excelList = new ArrayList<>();
            for (int i = 0; i < oriWorkbook.getNumberOfSheets(); i++) {
                XSSFSheet oriSheet = oriWorkbook.getSheetAt(i);
                XSSFSheet oriSameNameSheet = tarWorkbook.getSheet(oriSheet.getSheetName());

                if (oriSameNameSheet == null) {
                    CompareResult resExcel = new CompareResult(oriFile.getName());
                    resExcel.isSame = false;
                    resExcel.sheetName = oriSheet.getSheetName();
                    resExcel.errorType = CompareErrorType.MissingSheet;
                    resExcel.isTarget = false;
                    excelList.add(resExcel);
                    continue;
                }
                // *** 知道具体列数，可以先将oriSameNameSheet中的数据收集成map
                Map<Integer, String> keywordMapFromOriSameNameSheet =
                        collectKeywordFromSheet(oriSameNameSheet, keywordColList, dataStartLine);

                // *** 比较以第一次传入的excel为主体
                for (int k = dataStartLine; k < oriSheet.getPhysicalNumberOfRows(); k++) {
                    XSSFRow row = oriSheet.getRow(k);
                    if (row == null) {
                        break;
                    }
                    String keyword = genKeyword(row, keywordColList);
                    if (keyword.isEmpty()) {
                        break;
                    }
                    List<Integer> matchedKeyRowList = getMatchedKey(keyword, keywordMapFromOriSameNameSheet);
                    if (matchedKeyRowList.isEmpty()) {
                        // 关键字没匹配上
                        CompareResult resExcel = new CompareResult(oriFile.getName());
                        resExcel.isSame = false;
                        resExcel.sheetName = oriSheet.getSheetName();
                        resExcel.errorType = CompareErrorType.NotMatchedKeyword;
                        resExcel.rowId = k;
                        excelList.add(resExcel);
                    } else if (matchedKeyRowList.size() == 1) {
                        // 关键字匹配上
                        List<String> diffCols = compareRows(row, oriSameNameSheet.getRow(matchedKeyRowList.get(0)));
                        if (diffCols.isEmpty()) {
                            continue;
                        }
                        // 数据不一致
                        CompareResult resExcel = new CompareResult(oriFile.getName());
                        resExcel.isSame = false;
                        resExcel.sheetName = oriSheet.getSheetName();
                        resExcel.rowId = k;
                        resExcel.diffColumns = diffCols;
                        resExcel.errorType = CompareErrorType.DifferentValue;
                        excelList.add(resExcel);
                    } else {
                        // 比对行有多行
                        CompareResult resExcel = new CompareResult(oriFile.getName());
                        resExcel.isSame = false;
                        resExcel.sheetName = oriSheet.getSheetName();
                        resExcel.errorType = CompareErrorType.TooManyMatchedRows;
                        resExcel.rowId = k;
                        resExcel.matchedRowIds = matchedKeyRowList;
                        excelList.add(resExcel);
                    }
                }

                // *** 比较以第二次传入的excel为主体
                Map<Integer, String> keywordMapFromOriSheet =
                        collectKeywordFromSheet(oriSheet, keywordColList, dataStartLine);

                // 以file1为主体，比较file2中是否有多余的sheet完后再以file2为主体比较file1中是否有多余的sheet
                for (int m = dataStartLine; m < oriSameNameSheet.getPhysicalNumberOfRows(); m++) {
                    XSSFRow row = oriSameNameSheet.getRow(m);
                    if (row == null) {
                        break;
                    }
                    String keyword = genKeyword(row, keywordColList);
                    if (keyword.isEmpty()) {
                        break;
                    }
                    List<Integer> matchedKeyRowList = getMatchedKey(keyword, keywordMapFromOriSheet);
                    if (matchedKeyRowList.isEmpty()) {
                        // 关键字没匹配上
                        CompareResult resExcel = new CompareResult(tarFile.getName());
                        resExcel.isSame = false;
                        resExcel.sheetName = oriSameNameSheet.getSheetName();
                        resExcel.errorType = CompareErrorType.NotMatchedKeyword;
                        resExcel.rowId = m;
                        resExcel.isTarget = true;
                        excelList.add(resExcel);
                    } else if (matchedKeyRowList.size() > 1) {
                        // 匹配多行
                        CompareResult resExcel = new CompareResult(tarFile.getName());
                        resExcel.isSame = false;
                        resExcel.sheetName = oriSameNameSheet.getSheetName();
                        resExcel.errorType = CompareErrorType.TooManyMatchedRows;
                        resExcel.rowId = m;
                        resExcel.matchedRowIds = matchedKeyRowList;
                        excelList.add(resExcel);
                    }
                }
            }
            for (int j = 0; j < tarWorkbook.getNumberOfSheets(); j++){
                XSSFSheet tarSheet = tarWorkbook.getSheetAt(j);
                XSSFSheet tarSameNameSheet = oriWorkbook.getSheet(tarSheet.getSheetName());
                if (tarSameNameSheet == null ){
                    CompareResult resExcel = new CompareResult(oriFile.getName());
                    resExcel.isSame = false;
                    resExcel.sheetName = tarSheet.getSheetName();
                    resExcel.errorType = CompareErrorType.MissingSheet;
                    resExcel.isTarget = true;
                    excelList.add(resExcel);
                }
            }
            return excelList;
        }
    }

    // 比较两个行是否相等,并输出不相同的列
    public  List<String> compareRows(XSSFRow oriRow, XSSFRow tarRow) {
        List<String> diffCols = new ArrayList<>();
        int cells = Math.max(oriRow.getLastCellNum(), tarRow.getLastCellNum());
        for (int i = 0; i < cells; i++) {
            XSSFCell oriCell = oriRow.getCell(i);
            XSSFCell tarCell = tarRow.getCell(i);
            if (!compareCells(oriCell, tarCell)) {
                // 将第i列转换成具体列数
                diffCols.add(convertColToTitle(i + 1));
            }
        }
        return diffCols;
    }

    public  boolean compareCells(XSSFCell oriCell, XSSFCell tarCell) {
        if (oriCell == null && tarCell == null) {
            return true;
        } else if (oriCell == null || tarCell == null) {
            return false;
        }

        String value1 = oriCell.toString();
        String value2 = tarCell.toString();

        return value1.equals(value2);
    }

    public  void deleteFolder(File folder) {
        if (folder.exists()) {
            File[] files = folder.listFiles();
            if (files != null) {
                for (File file : files) {
                    if (file.isDirectory()) {
                        // 递归删除文件夹中的内容
                        deleteFolder(file);
                    } else {
                        // 删除文件
                        file.delete();
                    }
                }
            }
            folder.delete();
        }
    }

    public  String genKeyword(XSSFRow row, List<Integer> keyCols) {
        StringBuilder res = new StringBuilder();
        for (Integer keyCol : keyCols) {
            if (row.getCell(keyCol) != null) {
                XSSFCell cell = row.getCell(keyCol - 1);
                res.append(cell.toString());
            }
        }
        return res.toString();
    }

    public  Map<Integer, String> collectKeywordFromSheet(
            XSSFSheet sheet,
            List<Integer> keyCols,
            int dataStartLine
    ) {
        Map<Integer, String> map = new HashMap<>();
        for (int i = dataStartLine; i < sheet.getPhysicalNumberOfRows(); i++) {
            if (sheet.getRow(i) == null) {
                break;
            }
            XSSFRow row = sheet.getRow(i);
            String keyword = genKeyword(row, keyCols);
            map.put(i, keyword);
        }
        return map;
    }

    public  List<Integer> getMatchedKey(String keyword, Map<Integer, String> map) {
        List<Integer> list = new ArrayList<>();
        for (Map.Entry<Integer, String> entry : map.entrySet()) {
            int key = entry.getKey();
            String value = entry.getValue();
            if (keyword.equals(value)) {
                list.add(key);
            }
        }
        return list;
    }

    // 判断一个路径是否有文件夹，没有则创建
    public  void createFolder(String path) {
        File folder = new File(path);
        if (!folder.exists()) {
            folder.mkdir();
        } else {
            // 清空文件夹
            deleteFolder(folder);
            folder.mkdir();
        }
    }

    // 判断一个文件是不是以.xlsx或者.xls结尾
    public  boolean isExcelFile(File file) {
        return !file.getName().endsWith(".xlsx") && !file.getName().endsWith(".xls");
    }

    // 通过errorType往txt中写入信息
    public  void genErrorTxt(FileWriter writer, List<CompareResult> excelList) throws IOException {
        for (CompareResult excel : excelList) {
            if (excel != null && !excel.isSame) {
                String fileStr = "原始文件:[";
                if (excel.isTarget) {
                    fileStr = "对比文件:[";
                }
                switch (excel.errorType) {

                    case MissingFile:
                        writer.write(fileStr+ excel.fileName + "] 的比对文件不存在" + "\n");
                        break;
                    case MissingSheet:
                        writer.write(fileStr + excel.fileName + "]存在sheet：[" + excel.sheetName + "]，" +
                                "而另一压缩包中的此文件不存在该sheet" + "\n");
                        break;
                    case NotExcelFile:
                        writer.write(excel.fileName + " 不是excel文件" + "\n");
                        break;
                    case NotMatchedKeyword:
                        writer.write(fileStr+ excel.fileName + "]的sheet：[" + excel.sheetName + "] " +
                                "关键字没有匹配上，行号为第" + (excel.rowId + 1) + "行。" + "\n");
                        break;
                    case DifferentValue:
                        writer.write(fileStr+ excel.fileName + "]的sheet：[" +
                                excel.sheetName + "] 数据不一致，具体为第" + (excel.rowId + 1) + "行的" +
                                excel.diffColumns.toString()
                                        .replace("[", "").replace("]", "") +
                                "列" + "\n");
                        break;
                    case TooManyMatchedRows:
                        writer.write(fileStr+ excel.fileName + "]的sheet：[" + excel.sheetName
                                + "]比对行有多行,原始文件比对行为"
                                + (excel.rowId + 1) + "，对比文件的匹配行为" +
                                excel.matchedRowIds.toString().
                                        replace("[", "").replace("]", "") +
                                "列" + "\n");
                        break;
                }
            }
        }
    }

    // 加载配置
    public  Map<String, CompareFileConfig.FileConfig> loadConfig() throws IOException {
        ObjectMapper objectMapper = new ObjectMapper();
        Map<String, CompareFileConfig.FileConfig> map = new HashMap<>();
        // 读取用户想要的配置的excel_compare的配置文件
        ClassPathResource resource = new ClassPathResource("excel_compare.json");
        InputStream inputStream = resource.getStream();
        CompareFileConfig compareFile = objectMapper.readValue(inputStream, CompareFileConfig.class);
        List<CompareFileConfig.FileConfig> list = compareFile.getFiles();
        for (CompareFileConfig.FileConfig fileConfig : list) {
            map.put(fileConfig.fileName, fileConfig);
        }
        return map;
    }

    public  int columnToNumber(String column) {
        column = column.toUpperCase(); // 将列字母转换为大写

        int result = 0;
        for (int i = 0; i < column.length(); i++) {
            char c = column.charAt(i);
            int value = c - 'A' + 1; // 计算字母所代表的数字值
            result = result * 26 + value; // 将多个字母进行加权计算
        }

        return result;
    }

    public  String convertColToTitle(int n) {
        StringBuilder result = new StringBuilder();

        while (n > 0) {
            n--;
            char ch = (char) ('A' + n % 26);
            result.insert(0, ch);
            n /= 26;
        }

        return result.toString();
    }

    public void unzipMultipartFile(MultipartFile zipFile, String outputFolder) {
        byte[] buffer = new byte[1024];

        try {
            // 创建输出目录如果它不存在
            File folder = new File(outputFolder);
            if (!folder.exists()) {
                folder.mkdir();
            }

            // 获取ZIP文件的输入流
            InputStream zipInputStream = zipFile.getInputStream();
            ZipInputStream zis = new ZipInputStream(zipInputStream);
            ZipEntry ze = zis.getNextEntry();

            while (ze != null) {
                String fileName = ze.getName();
                File newFile = new File(outputFolder + File.separator + fileName);

                // 创建所有非存在的父目录
                new File(newFile.getParent()).mkdirs();

                if (ze.isDirectory()) {
                    newFile.mkdirs();
                } else {
                    // 写文件到磁盘
                    FileOutputStream fos = new FileOutputStream(newFile);
                    BufferedOutputStream bos = new BufferedOutputStream(fos);

                    int len;
                    while ((len = zis.read(buffer)) > 0) {
                        bos.write(buffer, 0, len);
                    }

                    bos.close();
                }
                ze = zis.getNextEntry();
            }

            zis.closeEntry();
            zis.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    public Resource loadFileAsResource(String url) {
        try {
            Path filePath = Paths.get(url).normalize();
            Resource resource = new UrlResource(filePath.toUri());
            if (resource.exists()) {
                return resource;
            } else {
                throw new FileNotFoundException("File not found ");
            }
        } catch (MalformedURLException e) {
            throw new RuntimeException(e);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
    }

    public static void zipFolder(String sourceFolder, String zipFilePath) throws IOException {
        FileOutputStream fos = new FileOutputStream(zipFilePath);
        ZipOutputStream zos = new ZipOutputStream(fos);
        addFolderToZip(sourceFolder, sourceFolder, zos);
        zos.close();
        fos.close();
    }

    private static void addFileToZip(String filePath, String parentFolder, ZipOutputStream zos) throws IOException {
        FileInputStream fis = new FileInputStream(filePath);

        String zipEntryName = filePath.substring(parentFolder.length() + 1);

        ZipEntry zipEntry = new ZipEntry(zipEntryName);
        zos.putNextEntry(zipEntry);

        byte[] buffer = new byte[1024];
        int length;
        while ((length = fis.read(buffer)) > 0) {
            zos.write(buffer, 0, length);
        }

        zos.closeEntry();
        fis.close();
    }

    private static void addFolderToZip(String sourceFolder, String parentFolder, ZipOutputStream zos) throws IOException {
        File folder = new File(sourceFolder);
        for (String fileName : folder.list()) {
            if (new File(sourceFolder + File.separator + fileName).isDirectory()) {
                addFolderToZip(sourceFolder + File.separator + fileName, parentFolder, zos);
            } else {
                addFileToZip(sourceFolder + File.separator + fileName, parentFolder, zos);
            }
        }
    }
}

