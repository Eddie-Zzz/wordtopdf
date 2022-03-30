package com.company.common.enums;

public enum FontEnum {

    SIM_SUN("SimSun","宋体", "D:\\Desktop\\OA开发相关资料\\内部签报\\字体文件\\simsun.ttf");

    // 字体名称
    private String id;
    // 字体中文名称
    private String chName;
    // 字体文件路径
    private String path;

    private FontEnum(String id, String name, String path) {
        this.id = id;
        this.path = path;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getChName() {
        return chName;
    }

    public void setChName(String chName) {
        this.chName = chName;
    }

    public String getPath() {
        return path;
    }

    public void setPath(String path) {
        this.path = path;
    }
}
