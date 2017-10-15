package com.nouseen.bean;

/**
 * Created by nouseen on 2017/9/20.
 */
public class Ant {
    // 编号
    private long antID;

    // 速度
    private double speed;

    // 位置
    private double position;

    // 方向
    private int headTo;

    // 离场时间
    private int outTime;

    public int getOutTime() {
        return outTime;
    }

    public void setOutTime(int outTime) {
        this.outTime = outTime;
    }

    public Ant(){
        super();
    }

    public Ant(long antID, double speed, double position, int headTo) {
        this.antID = antID;
        this.speed = speed;
        this.position = position;
        this.headTo = headTo;
    }

    public long getAntID() {
        return antID;
    }

    public void setAntID(long antID) {
        this.antID = antID;
    }

    public double getSpeed() {
        return speed;
    }

    public void setSpeed(double speed) {
        this.speed = speed;
    }

    public double getPosition() {
        return position;
    }

    public void setPosition(double position) {
        this.position = position;
    }

    public int getHeadTo() {
        return headTo;
    }

    public void setHeadTo(int headTo) {
        this.headTo = headTo;
    }
}
