package org.example.compare.controller;

import org.example.compare.bean.CompareFileConfig;
import org.example.compare.bean.FileResult;
import org.example.util;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Map;

@Controller
@CrossOrigin(origins = "http://localhost:8000", maxAge = 3600)
@RequestMapping("/excel")
public class ExcelCompareController {
    @Autowired
    public ExcelComparator excelComparator;

    final String globalDir = "/Users/houjuntao/test/";
    @PostMapping("/compare")
    @ResponseBody
    public List<FileResult> compare(@RequestParam("oriZip") MultipartFile oriZip, @RequestParam("tarZip") MultipartFile tarZip){
        try {
            // 读取excel_compare.json文件
            Map<String, CompareFileConfig.FileConfig> configMap = excelComparator.loadConfig();

            // 生成一个随机时间戳
            long timestamp = System.currentTimeMillis();
            String compareFileDir1 = globalDir + "compareFileDir1/" + timestamp;
            String compareFileDir2 = globalDir + "compareFileDir2/" + timestamp;
            // 判断下是否存在临时文件夹，不存在则创建
            excelComparator.createFolder(compareFileDir1);
            excelComparator.createFolder(compareFileDir2);

            // 解压两个zip文件
            excelComparator.unzipMultipartFile(oriZip, compareFileDir1);
            excelComparator.unzipMultipartFile(tarZip, compareFileDir2);

            // 遍历解压后的文件夹，比较文件
            Map<String, List<CompareResult>> resMap = excelComparator.compareExcelFiles(compareFileDir1, compareFileDir2, configMap);
            // 检查输出文件目录是否存在，如果不存在，则创建它
            SimpleDateFormat timeFormat = new SimpleDateFormat("yyyyMMddHHmmss");
            String timeStr = timeFormat.format(System.currentTimeMillis());
            String outputTxtPath = globalDir + "outputTxt/"+timeStr;
            File outputTxtFloder = new File(outputTxtPath);
            if (!outputTxtFloder.exists()) {
                boolean dirCreated = outputTxtFloder.mkdirs(); // 创建目录及其所有父目录
                if (!dirCreated) {
                    throw new RuntimeException("Failed to create output directory");
                }
            }
            // 将resMap遍历，每个文件生成一个txt文件
            List<FileResult> fileResults = new ArrayList<>();
            for (Map.Entry<String, List<CompareResult>> entry : resMap.entrySet()) {
                String outputFilePath = globalDir + "outputTxt/"+timeStr + "/" + entry.getKey() + ".txt";
                try (FileWriter writer = new FileWriter(outputFilePath)) {
                    excelComparator.genErrorTxt(writer,entry.getValue());
                }
                FileResult fileResult = new FileResult();
                Path path = Paths.get(outputFilePath);
                String content = Files.readString(path);
                fileResult.setFilePath(outputFilePath);
                fileResult.setContent(content);
                fileResult.setFileName(util.removeAfterLastDot(entry.getKey()));
                fileResults.add(fileResult);
            }
            // 将文件夹打包成zip文件
            String outputZipDir = globalDir + "outputZip/";
            String outputZipPath = globalDir + "outputZip/"+timeStr + ".zip";
            File outputZipFloder = new File(outputZipDir);
            if (!outputZipFloder.exists()) {
                boolean dirCreated = outputZipFloder.mkdirs(); // 创建目录及其所有父目录
                if (!dirCreated) {
                    throw new RuntimeException("Failed to create output directory");
                }
            }
            excelComparator.zipFolder(outputTxtPath, outputZipPath);
            // 压缩完删除outputTxt文件夹
            // 删除临时文件夹
            excelComparator.deleteFolder(new File(compareFileDir1));
            excelComparator.deleteFolder(new File(compareFileDir2));
            excelComparator.deleteFolder(new File(outputTxtPath));
            // 将fileResults按照文件名排序
            fileResults.sort(Comparator.comparing(FileResult::getFileName));
            return fileResults;
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    @PostMapping("/compare/download")
    @ResponseBody
    public ResponseEntity<Resource> downloadFile(@RequestParam("filePath") String url) {
        Resource resource = excelComparator.loadFileAsResource(url);
        String contentType = "";
        try {
            contentType = Files.probeContentType(Paths.get(resource.getURI()));
        } catch (IOException e) {
            contentType = MediaType.APPLICATION_OCTET_STREAM_VALUE;
        }

        return ResponseEntity.ok()
                .contentType(MediaType.parseMediaType(contentType))
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + resource.getFilename() + "\"")
                .body(resource);
    }
}

