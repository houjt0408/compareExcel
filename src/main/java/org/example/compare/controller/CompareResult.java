package org.example.compare.controller;

import java.util.List;

public class CompareResult {
    String fileName;
    String sheetName;
    Integer rowId;
    List<Integer> matchedRowIds;
    List<String> diffColumns;
    boolean isSame;

    boolean isTarget;
    CompareErrorType errorType;

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public Integer getRowId() {
        return rowId;
    }

    public void setRowId(Integer rowId) {
        this.rowId = rowId;
    }

    public List<Integer> getMatchedRowIds() {
        return matchedRowIds;
    }

    public void setMatchedRowIds(List<Integer> matchedRowIds) {
        this.matchedRowIds = matchedRowIds;
    }

    public List<String> getDiffColumns() {
        return diffColumns;
    }

    public void setDiffColumns(List<String> diffColumns) {
        this.diffColumns = diffColumns;
    }

    public boolean isSame() {
        return isSame;
    }

    public void setSame(boolean same) {
        isSame = same;
    }

    public boolean isTarget() {
        return isTarget;
    }

    public void setTarget(boolean target) {
        isTarget = target;
    }

    public CompareErrorType getErrorType() {
        return errorType;
    }

    public void setErrorType(CompareErrorType errorType) {
        this.errorType = errorType;
    }

    public CompareResult(String name) {
        this.fileName = name;
        this.sheetName = "";
        this.rowId = 0;
        this.matchedRowIds = null;
        this.diffColumns = null;
        this.isSame = true;
        this.errorType = null;
        this.isTarget = false;
    }
}

enum CompareErrorType {
    MissingFile,
    MissingSheet,
    DifferentValue,
    NotExcelFile,
    // 关键字没有匹配上
    NotMatchedKeyword,
    // 比对行有多行
    TooManyMatchedRows,
}