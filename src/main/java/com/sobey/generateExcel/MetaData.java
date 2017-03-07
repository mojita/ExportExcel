package com.sobey.generateExcel;

/**
 * Created by lijunhong on 17/3/6.
 */
public class MetaData {

    private String fileId;              //文件序号
    private String title;               //题目
    private String secretLevel;          //保密等级
    private String fileState;           //文件状态
    private String pageNum;             //页数
    private String fileNo;              //档号
    private String videoLength;         //片长
    private String shootingDate;        //拍摄时间
    private String shootingAddress;     //拍摄地址
    private String shootingUnit;        //拍摄单位
    private String photographer;        //拍摄者


    public String getFileId() {
        return fileId;
    }

    public void setFileId(String fileId) {
        this.fileId = fileId;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String getSecretLevel() {
        return secretLevel;
    }

    public void setSecretLevel(String secretLevel) {
        this.secretLevel = secretLevel;
    }

    public String getFileState() {
        return fileState;
    }

    public void setFileState(String fileState) {
        this.fileState = fileState;
    }

    public String getPageNum() {
        return pageNum;
    }

    public void setPageNum(String pageNum) {
        this.pageNum = pageNum;
    }

    public String getFileNo() {
        return fileNo;
    }

    public void setFileNo(String fileNo) {
        this.fileNo = fileNo;
    }

    public String getVideoLength() {
        return videoLength;
    }

    public void setVideoLength(String videoLength) {
        this.videoLength = videoLength;
    }

    public String getShootingDate() {
        return shootingDate;
    }

    public void setShootingDate(String shootingDate) {
        this.shootingDate = shootingDate;
    }

    public String getShootingAddress() {
        return shootingAddress;
    }

    public void setShootingAddress(String shootingAddress) {
        this.shootingAddress = shootingAddress;
    }

    public String getShootingUnit() {
        return shootingUnit;
    }

    public void setShootingUnit(String shootingUnit) {
        this.shootingUnit = shootingUnit;
    }

    public String getPhotographer() {
        return photographer;
    }

    public void setPhotographer(String photographer) {
        this.photographer = photographer;
    }


    @Override
    public String toString() {
        return "MetaData{" +
                "fileId='" + fileId + '\'' +
                ", title='" + title + '\'' +
                ", secretLevel='" + secretLevel + '\'' +
                ", fileState='" + fileState + '\'' +
                ", pageNum='" + pageNum + '\'' +
                ", fileNo='" + fileNo + '\'' +
                ", videoLength='" + videoLength + '\'' +
                ", shootingDate='" + shootingDate + '\'' +
                ", shootingAddress='" + shootingAddress + '\'' +
                ", shootingUnit='" + shootingUnit + '\'' +
                ", photographer='" + photographer + '\'' +
                '}';
    }
}
